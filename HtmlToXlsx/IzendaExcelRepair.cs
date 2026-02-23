
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using HtmlAgilityPack;

namespace GraspBI.Izenda
{
    public static class IzendaExcelRepair
    {
        private static readonly Dictionary<string, string> FormatMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["xls-text"] = "@",
            ["xls-l-text"] = "General",
            ["xls-percent"] = "0%",
            ["xls-date"] = "Short Date",
            ["xls-time"] = "h:mm AM/PM",
            ["xls-date-ShortDateFormat"] = "M/d/yyyy",
            ["xls-date-LongDateFormat"] = "dddd, MMMM d, yyyy",
            ["xls-date-ShortTimeFormat"] = "h:mm AM/PM",
            ["xls-date-LongTimeFormat"] = "h:mm:ss AM/PM",
            ["xls-date-FullShortFormat"] = "dddd, MMMM d, yyyy h:mm AM/PM",
            ["xls-date-FullLongFormat"] = "dddd, MMMM d, yyyy h:mm:ss AM/PM",
            ["xls-date-GenShortFormat"] = "M/d/yyyy h:mm AM/PM",
            ["xls-date-GenLongFormat"] = "M/d/yyyy h:mm:ss AM/PM",
        };

        private static readonly Dictionary<string, RowStyle> DefaultRowStyles = new Dictionary<string, RowStyle>(StringComparer.OrdinalIgnoreCase)
        {
            ["ReportHeader"] = new RowStyle(XLColor.White, XLColor.Black, true, 9),
            ["ReportItem"] = new RowStyle(XLColor.Black, XLColor.White, false, 8),
            ["AlternatingItem"] = new RowStyle(XLColor.Black, XLColor.FromName("Gainsboro"), false, 8),
            ["ReportFooter"] = new RowStyle(XLColor.Black, XLColor.White, true, 8),
        };

        private static readonly TableStyle DefaultTableStyle = new TableStyle(XLColor.Black);

        public static void RepairFile(string inputPath, string outputPath)
        {
            var html = File.ReadAllText(inputPath);
            var mimeImages = ExtractMimeImages(html);

            var doc = new HtmlDocument();
            doc.LoadHtml(html);

            Dictionary<string, RowStyle> rowStyles;
            TableStyle tableStyle;
            ParseCssStyles(doc, out rowStyles, out tableStyle);

            var titleNode = doc.DocumentNode.SelectSingleNode("//span[contains(@class,'ReportTitle')]");
            var descNode = doc.DocumentNode.SelectSingleNode("//span[contains(@class,'Description')]");
            var reportTitle = titleNode != null ? (titleNode.InnerText ?? "").Trim() : null;
            var reportDesc = descNode != null ? (descNode.InnerText ?? "").Trim() : null;

            // Find header logo (inside the header table, separate from chart images)
            var headerLogoImages = FindHeaderLogoImages(doc);

            // Find report sections in document order via report= attribute
            var sectionDivs = doc.DocumentNode.SelectNodes("//div[@report]");

            using (var wb = new XLWorkbook())
            {
                var sheetName = !string.IsNullOrWhiteSpace(reportTitle) && reportTitle.Length <= 31
                    ? SanitizeSheetName(reportTitle)
                    : "Report";
                var ws = wb.AddWorksheet(sheetName);

                int startRow = 1;
                int imageIndex = 0;
                startRow = InsertImages(ws, headerLogoImages, mimeImages, startRow, ref imageIndex);

                if (!string.IsNullOrWhiteSpace(reportTitle))
                {
                    var cell = ws.Cell(startRow, 1);
                    cell.Value = reportTitle;
                    cell.Style.Font.Bold = true;
                    cell.Style.Font.FontSize = 16;
                    startRow++;
                }

                if (!string.IsNullOrWhiteSpace(reportDesc))
                {
                    var cell = ws.Cell(startRow, 1);
                    cell.Value = reportDesc;
                    cell.Style.Font.FontSize = 10;
                    cell.Style.Font.FontColor = XLColor.Gray;
                    startRow++;
                }

                if (startRow > 1)
                    startRow++;

                // Process each report section in document order
                if (sectionDivs != null)
                {
                    for (int sectionIdx = 0; sectionIdx < sectionDivs.Count; sectionIdx++)
                    {
                        var sectionDiv = sectionDivs[sectionIdx];
                        var reportAttr = sectionDiv.GetAttributeValue("report", "");

                        if (reportAttr.StartsWith("Chart", StringComparison.OrdinalIgnoreCase))
                        {
                            // Chart section: insert chart images
                            var chartImages = FindCidImagesInNode(sectionDiv);
                            startRow = InsertImages(ws, chartImages, mimeImages, startRow, ref imageIndex);
                        }
                        else
                        {
                            // Data section (Detail, Summary, etc.): insert ReportTable
                            var reportTable = sectionDiv.SelectSingleNode(".//table[contains(@class,'ReportTable')]");
                            if (reportTable == null) continue;

                            var rows = reportTable.SelectNodes(".//tr");
                            if (rows == null || rows.Count == 0) continue;

                            if (sectionIdx > 0 && startRow > 1)
                                startRow++;

                            for (int r = 0; r < rows.Count; r++)
                            {
                                var tr = rows[r];
                                var rowClass = GetClass(tr);
                                var cells = tr.SelectNodes("td|th");
                                if (cells == null) continue;

                                int col = 1;
                                for (int c = 0; c < cells.Count; c++)
                                {
                                    var td = cells[c];
                                    int colspan = td.GetAttributeValue("colspan", 1);
                                    if (colspan < 1) colspan = 1;

                                    var xlCell = ws.Cell(startRow + r, col);
                                    var rawText = HtmlEntity.DeEntitize(td.InnerText).Trim();
                                    var formatClass = DetectFormatClass(td);
                                    SetCellValue(xlCell, rawText, formatClass);
                                    ApplyAlignment(xlCell, td);
                                    var cellClass = GetClass(td);
                                    var effectiveClass = !string.IsNullOrEmpty(rowClass) ? rowClass : cellClass;
                                    ApplyRowStyle(xlCell, effectiveClass, rowStyles, tableStyle);

                                    if (colspan > 1)
                                    {
                                        var mergeRange = ws.Range(startRow + r, col, startRow + r, col + colspan - 1);
                                        mergeRange.Merge();
                                        foreach (var mergedCell in mergeRange.Cells())
                                        {
                                            if (mergedCell != xlCell)
                                                ApplyRowStyle(mergedCell, effectiveClass, rowStyles, tableStyle);
                                        }
                                    }

                                    col += colspan;
                                }
                            }

                            startRow += rows.Count;
                            startRow++; // gap after table
                        }
                    }
                }

                ws.Columns().AdjustToContents();

                foreach (var col in ws.ColumnsUsed())
                {
                    if (col.Width > 50)
                        col.Width = 50;
                }

                wb.SaveAs(outputPath);
            }
        }

        private static string DetectFormatClass(HtmlNode td)
        {
            var tdClass = GetClass(td);
            var match = FindFormatClass(tdClass);
            if (match != null) return match;

            var childDivs = td.SelectNodes(".//div[@class]");
            if (childDivs == null) return null;

            return childDivs
                .Select(div => FindFormatClass(GetClass(div)))
                .FirstOrDefault(m => m != null);
        }

        private static string FindFormatClass(string cssClasses)
        {
            if (string.IsNullOrEmpty(cssClasses)) return null;

            return FormatMap.Keys
                .OrderByDescending(k => k.Length)
                .FirstOrDefault(key => cssClasses.IndexOf(key, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private static void SetCellValue(IXLCell cell, string rawText, string formatClass)
        {
            if (string.IsNullOrWhiteSpace(rawText) || rawText == "&nbsp;" || rawText == "\u00A0")
            {
                cell.Value = "";
                return;
            }

            if (formatClass != null && (formatClass.StartsWith("xls-date", StringComparison.OrdinalIgnoreCase)
                || formatClass == "xls-time"))
            {
                double serial;
                if (double.TryParse(rawText, NumberStyles.Float, CultureInfo.InvariantCulture, out serial))
                {
                    try
                    {
                        cell.Value = DateTime.FromOADate(serial);
                        cell.Style.NumberFormat.Format = FormatMap[formatClass];
                        return;
                    }
                    catch
                    {
                    }
                }
                cell.Value = rawText;
                cell.Style.NumberFormat.Format = "@";
                return;
            }

            if (formatClass == "xls-percent")
            {
                double pct;
                if (double.TryParse(rawText.TrimEnd('%'), NumberStyles.Float, CultureInfo.InvariantCulture, out pct))
                {
                    cell.Value = pct / 100.0;
                    cell.Style.NumberFormat.Format = FormatMap["xls-percent"];
                    return;
                }
            }

            if (formatClass == "xls-text" || formatClass == "xls-l-text")
            {
                cell.SetValue(rawText);
                cell.Style.NumberFormat.Format = "@";
                return;
            }

            decimal amount;
            if (TryParseCurrency(rawText, out amount))
            {
                cell.Value = (double)amount;
                cell.Style.NumberFormat.Format = "$#,##0.00";
                return;
            }

            double num;
            if (double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands,
                                CultureInfo.InvariantCulture, out num))
            {
                cell.Value = num;
                return;
            }

            cell.Value = rawText;
        }

        private static bool TryParseCurrency(string text, out decimal amount)
        {
            amount = 0;
            if (!text.Contains("$")) return false;

            var cleaned = text.Replace("$", "").Replace(",", "").Trim();
            return decimal.TryParse(cleaned, NumberStyles.Float | NumberStyles.AllowLeadingSign,
                                    CultureInfo.InvariantCulture, out amount);
        }

        private static void ApplyAlignment(IXLCell cell, HtmlNode td)
        {
            var align = td.GetAttributeValue("align", "left").ToLowerInvariant();
            switch (align)
            {
                case "center":
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    break;
                case "right":
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                    break;
                default:
                    cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                    break;
            }
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
        }

        private static void ApplyRowStyle(IXLCell cell, string rowClass,
                                   Dictionary<string, RowStyle> rowStyles, TableStyle tableStyle)
        {
            var matchingStyle = rowStyles.FirstOrDefault(kvp =>
                rowClass.IndexOf(kvp.Key, StringComparison.OrdinalIgnoreCase) >= 0);

            if (matchingStyle.Key != null)
            {
                var style = matchingStyle.Value;
                cell.Style.Font.FontColor = style.FontColor;
                cell.Style.Fill.BackgroundColor = style.BackgroundColor;
                cell.Style.Font.Bold = style.Bold;
                cell.Style.Font.Italic = style.Italic;
                cell.Style.Font.FontSize = style.FontSize;

                if (matchingStyle.Key == "ReportFooter")
                {
                    cell.Style.Border.TopBorder = XLBorderStyleValues.Double;
                    cell.Style.Border.TopBorderColor = tableStyle.BorderColor;
                }
            }

            // Always apply borders and font to all cells in the table
            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            cell.Style.Border.OutsideBorderColor = tableStyle.BorderColor;
            cell.Style.Font.FontName = "Tahoma";
        }

        private static string GetClass(HtmlNode node)
        {
            return node.GetAttributeValue("class", "");
        }

        private static string SanitizeSheetName(string name)
        {
            var sanitized = Regex.Replace(name, @"[\\/?*\[\]:]", "_");
            return sanitized.Length > 31 ? sanitized.Substring(0, 31) : sanitized;
        }

        private static void ParseCssStyles(HtmlDocument doc, out Dictionary<string, RowStyle> rowStyles, out TableStyle tableStyle)
        {
            var styleNodes = doc.DocumentNode.SelectNodes("//style");
            if (styleNodes == null || styleNodes.Count == 0)
            {
                rowStyles = new Dictionary<string, RowStyle>(DefaultRowStyles, StringComparer.OrdinalIgnoreCase);
                tableStyle = DefaultTableStyle;
                return;
            }

            var cssText = string.Join("\n", styleNodes.Select(s => s.InnerText));

            var result = new Dictionary<string, RowStyle>(StringComparer.OrdinalIgnoreCase);
            foreach (var className in new[] { "ReportHeader", "ReportItem", "AlternatingItem", "ReportFooter" })
            {
                RowStyle def;
                if (!DefaultRowStyles.TryGetValue(className, out def))
                    def = new RowStyle(XLColor.Black, XLColor.White, false, 8);

                var fontColorStr = FindCssProperty(cssText, className, "color");
                var bgColorStr = FindCssProperty(cssText, className, "background-color");
                var fontWeightStr = FindCssProperty(cssText, className, "font-weight");
                var fontStyleStr = FindCssProperty(cssText, className, "font-style");
                var fontSizeStr = FindCssProperty(cssText, className, "font-size");

                var fontColor = fontColorStr != null ? ParseCssColor(fontColorStr) : def.FontColor;
                var bgColor = bgColorStr != null ? ParseCssColor(bgColorStr) : def.BackgroundColor;
                var bold = fontWeightStr != null
                    ? fontWeightStr.Equals("bold", StringComparison.OrdinalIgnoreCase)
                    : def.Bold;
                var italic = fontStyleStr != null
                    && fontStyleStr.Equals("italic", StringComparison.OrdinalIgnoreCase);
                var fontSize = fontSizeStr != null ? ParseFontSize(fontSizeStr) : def.FontSize;

                result[className] = new RowStyle(fontColor, bgColor, bold, fontSize, italic);
            }

            var tableBorderColorStr = FindCssProperty(cssText, "ReportTable", "border-color");
            tableStyle = tableBorderColorStr != null
                ? new TableStyle(ParseCssColor(tableBorderColorStr))
                : DefaultTableStyle;

            var cellBorderColor = FindCellBorderColor(cssText);
            if (cellBorderColor != null)
                tableStyle = new TableStyle(cellBorderColor);

            rowStyles = result;
        }

        private static string FindCssProperty(string css, string className, string property)
        {
            var blockPattern = new Regex(
                @"[^{}]*\." + Regex.Escape(className) + @"\b[^{}]*\{(?<props>[^}]*)\}",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);

            string lastValue = null;
            foreach (Match block in blockPattern.Matches(css))
            {
                var props = block.Groups["props"].Value;

                string prefix = property == "color" ? @"(?<![a-zA-Z-])" : "";
                string propRegex = prefix + Regex.Escape(property) + @"\s*:\s*(?<value>[^;}\n]+?)\s*(?:[;}\n]|$)";

                var propMatch = Regex.Match(props, propRegex, RegexOptions.IgnoreCase);
                if (propMatch.Success)
                {
                    var val = propMatch.Groups["value"].Value.Trim();
                    if (!val.Equals("inherit", StringComparison.OrdinalIgnoreCase) &&
                        !val.Equals("transparent", StringComparison.OrdinalIgnoreCase) &&
                        !val.Equals("initial", StringComparison.OrdinalIgnoreCase))
                    {
                        lastValue = val;
                    }
                }
            }

            return lastValue;
        }

        private static XLColor FindCellBorderColor(string css)
        {
            var pattern = new Regex(
                @"(?:ReportHeader|ReportItem|AlternatingItem)\s[^{}]*\{[^}]*\bborder\s*:\s*(?<value>[^;}\n]+)",
                RegexOptions.IgnoreCase);

            var match = pattern.Match(css);
            if (!match.Success) return null;

            var value = match.Groups["value"].Value.Trim();

            var hexMatch = Regex.Match(value, @"(#[0-9a-fA-F]{3,8})");
            if (hexMatch.Success)
                return ParseCssColor(hexMatch.Groups[1].Value);

            var colorToken = value.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
                .Reverse()
                .FirstOrDefault(t => !Regex.IsMatch(t, @"^\d") &&
                    !Regex.IsMatch(t, @"^(solid|dashed|dotted|double|none|groove|ridge|inset|outset)$", RegexOptions.IgnoreCase));

            return colorToken != null ? ParseCssColor(colorToken) : null;
        }

        private static XLColor ParseCssColor(string cssColor)
        {
            cssColor = cssColor.Trim();
            if (cssColor.StartsWith("#")) return XLColor.FromHtml(cssColor);

            switch (cssColor.ToLowerInvariant())
            {
                case "black": return XLColor.Black;
                case "white": return XLColor.White;
                case "red": return XLColor.Red;
                case "blue": return XLColor.Blue;
                case "green": return XLColor.Green;
                case "yellow": return XLColor.Yellow;
                case "gainsboro": return XLColor.FromName("Gainsboro");
                case "gray":
                case "grey": return XLColor.Gray;
                case "silver": return XLColor.FromName("Silver");
                case "transparent":
                case "inherit": return XLColor.NoColor;
                default: return TryParseNamedColor(cssColor);
            }
        }

        private static XLColor TryParseNamedColor(string name)
        {
            try
            {
                return XLColor.FromName(name);
            }
            catch
            {
                return XLColor.Black;
            }
        }

        private static double ParseFontSize(string cssFontSize)
        {
            var match = Regex.Match(cssFontSize, @"([\d.]+)");
            double size;
            return match.Success && double.TryParse(match.Groups[1].Value,
                NumberStyles.Float, CultureInfo.InvariantCulture, out size)
                ? size : 8;
        }

        private static Dictionary<string, byte[]> ExtractMimeImages(string rawContent)
        {
            var images = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);

            var htmlEnd = rawContent.IndexOf("</html>", StringComparison.OrdinalIgnoreCase);
            if (htmlEnd < 0) return images;

            var mimeSection = rawContent.Substring(htmlEnd + "</html>".Length);

            var mimePartPattern = new Regex(
                @"Content-ID:\s*(?<cid>\S+).*?Content-Transfer-Encoding:\s*BASE64\s*\r?\n\s*\r?\n(?<data>[A-Za-z0-9+/=\s]+?)(?=\r?\n--|\s*$)",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);

            foreach (Match m in mimePartPattern.Matches(mimeSection))
            {
                var cid = m.Groups["cid"].Value.Trim();
                var base64 = m.Groups["data"].Value.Trim();
                base64 = Regex.Replace(base64, @"\s+", "");

                try
                {
                    var bytes = System.Convert.FromBase64String(base64);
                    images[cid] = bytes;
                }
                catch (FormatException)
                {
                }
            }

            return images;
        }
       
        private static List<CidImageRef> FindHeaderLogoImages(HtmlDocument doc)
        {
            // Logo is inside the header table (the one containing the ReportTitle), not in any report section
            var headerTable = doc.DocumentNode.SelectSingleNode("//table[.//span[contains(@class,'ReportTitle')]]");
            if (headerTable == null) return new List<CidImageRef>();
            return FindCidImagesInNode(headerTable);
        }

        private static List<CidImageRef> FindCidImagesInNode(HtmlNode node)
        {
            var result = new List<CidImageRef>();

            var imgNodes = node.SelectNodes(".//img[contains(@src,'cid:')]");
            if (imgNodes == null) return result;

            foreach (var img in imgNodes)
            {
                var src = img.GetAttributeValue("src", "");
                if (!src.StartsWith("cid:", StringComparison.OrdinalIgnoreCase)) continue;

                var cid = src.Substring("cid:".Length).Trim();
                var width = img.GetAttributeValue("width", 0);
                var height = img.GetAttributeValue("height", 0);

                if (width == 0 || height == 0)
                {
                    var style = img.GetAttributeValue("style", "");
                    if (!string.IsNullOrEmpty(style))
                    {
                        var wMatch = Regex.Match(style, @"width:\s*(\d+)px", RegexOptions.IgnoreCase);
                        var hMatch = Regex.Match(style, @"height:\s*(\d+)px", RegexOptions.IgnoreCase);
                        if (wMatch.Success) width = int.Parse(wMatch.Groups[1].Value);
                        if (hMatch.Success) height = int.Parse(hMatch.Groups[1].Value);
                    }
                }

                result.Add(new CidImageRef(cid, width, height));
            }

            return result;
        }

        private static int InsertImages(IXLWorksheet ws, List<CidImageRef> imgRefs, Dictionary<string, byte[]> mimeImages, int startRow, ref int imageIndex)
        {
            if (imgRefs.Count == 0 || mimeImages.Count == 0)
                return startRow;

            foreach (var imgRef in imgRefs)
            {
                byte[] imageBytes;
                if (!mimeImages.TryGetValue(imgRef.ContentId, out imageBytes))
                    continue;

                imageIndex++;
                using (var ms = new MemoryStream(imageBytes))
                {
                    var picture = ws.AddPicture(ms, string.Format("Image_{0}", imageIndex))
                        .MoveTo(ws.Cell(startRow, 1));

                    if (imgRef.Width > 0 && imgRef.Height > 0)
                    {
                        const int maxWidthPx = 750;
                        if (imgRef.Width > maxWidthPx)
                        {
                            double scale = (double)maxWidthPx / imgRef.Width;
                            picture.ScaleWidth(scale);
                            picture.ScaleHeight(scale);
                        }
                    }
                    else
                    {
                        picture.Scale(0.5);
                    }
                }

                var imgHeight = imgRef.Height > 0 ? imgRef.Height : 200;
                const int maxW = 750;
                if (imgRef.Width > maxW && imgRef.Width > 0)
                    imgHeight = (int)(imgHeight * ((double)maxW / imgRef.Width));
                int rowsForImage = Math.Max(1, (int)Math.Ceiling(imgHeight / 15.0));

                startRow += rowsForImage + 1;
            }

            return startRow;
        }
    }

    public sealed class RowStyle
    {
        public XLColor FontColor { get; }
        public XLColor BackgroundColor { get; }
        public bool Bold { get; }
        public double FontSize { get; }
        public bool Italic { get; }

        public RowStyle(XLColor fontColor, XLColor backgroundColor, bool bold,
                        double fontSize = 8, bool italic = false)
        {
            FontColor = fontColor;
            BackgroundColor = backgroundColor;
            Bold = bold;
            FontSize = fontSize;
            Italic = italic;
        }
    }

    public sealed class CidImageRef
    {
        public string ContentId { get; }
        public int Width { get; }
        public int Height { get; }

        public CidImageRef(string contentId, int width, int height)
        {
            ContentId = contentId;
            Width = width;
            Height = height;
        }
    }

    public sealed class TableStyle
    {
        public XLColor BorderColor { get; }

        public TableStyle(XLColor borderColor)
        {
            BorderColor = borderColor;
        }
    }
}