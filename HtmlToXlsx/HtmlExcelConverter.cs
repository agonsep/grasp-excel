using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using HtmlAgilityPack;

namespace HtmlToXlsx;

/// <summary>
/// Converts HTML-based Excel reports (produced by Izenda/SSRS-style tools) into proper .xlsx files.
/// Parses CSS classes for Excel number formats, row styling, cell alignment, and embedded MIME images.
/// </summary>
public class HtmlExcelConverter
{
    // ----- Excel number format map (CSS class → Excel format string) -----
    private static readonly Dictionary<string, string> FormatMap = new(StringComparer.OrdinalIgnoreCase)
    {
        ["xls-text"]                 = "@",
        ["xls-l-text"]               = "General",
        ["xls-percent"]              = "0%",
        ["xls-date"]                 = "Short Date",
        ["xls-time"]                 = "h:mm AM/PM",
        ["xls-date-ShortDateFormat"] = "M/d/yyyy",
        ["xls-date-LongDateFormat"]  = "dddd, MMMM d, yyyy",
        ["xls-date-ShortTimeFormat"] = "h:mm AM/PM",
        ["xls-date-LongTimeFormat"]  = "h:mm:ss AM/PM",
        ["xls-date-FullShortFormat"] = "dddd, MMMM d, yyyy h:mm AM/PM",
        ["xls-date-FullLongFormat"]  = "dddd, MMMM d, yyyy h:mm:ss AM/PM",
        ["xls-date-GenShortFormat"]  = "M/d/yyyy h:mm AM/PM",
        ["xls-date-GenLongFormat"]   = "M/d/yyyy h:mm:ss AM/PM",
    };

    // ----- Row style definitions (CSS class → style info) -----
    private static readonly Dictionary<string, RowStyle> RowStyles = new(StringComparer.OrdinalIgnoreCase)
    {
        ["ReportHeader"]     = new RowStyle(XLColor.White, XLColor.Black, true),
        ["ReportItem"]       = new RowStyle(XLColor.Black, XLColor.White, false),
        ["AlternatingItem"]  = new RowStyle(XLColor.Black, XLColor.FromName("Gainsboro"), false),
        ["ReportFooter"]     = new RowStyle(XLColor.Black, XLColor.White, true),
    };

    /// <summary>
    /// Convert an HTML-based .xls/.mht file to a proper .xlsx file.
    /// </summary>
    public void Convert(string inputPath, string outputPath)
    {
        var html = File.ReadAllText(inputPath);

        // Extract embedded MIME images (Content-ID → image bytes)
        var mimeImages = ExtractMimeImages(html);

        var doc = new HtmlDocument();
        doc.LoadHtml(html);

        // Find the main data table (class="ReportTable")
        var reportTable = doc.DocumentNode.SelectSingleNode("//table[contains(@class,'ReportTable')]");
        if (reportTable == null)
            throw new InvalidOperationException("No <table class='ReportTable'> found in the input file.");

        // Extract report title and description if present
        var titleNode = doc.DocumentNode.SelectSingleNode("//span[contains(@class,'ReportTitle')]");
        var descNode = doc.DocumentNode.SelectSingleNode("//span[contains(@class,'Description')]");
        var reportTitle = titleNode?.InnerText?.Trim();
        var reportDesc = descNode?.InnerText?.Trim();

        // Find all <img> tags with cid: references
        var imgNodes = FindCidImages(doc);

        // Get all rows
        var rows = reportTable.SelectNodes(".//tr");
        if (rows == null || rows.Count == 0)
            throw new InvalidOperationException("No rows found in the report table.");

        using var wb = new XLWorkbook();
        var sheetName = !string.IsNullOrWhiteSpace(reportTitle) && reportTitle.Length <= 31
            ? SanitizeSheetName(reportTitle)
            : "Report";
        var ws = wb.AddWorksheet(sheetName);

        int startRow = 1;

        // Insert embedded images at the top of the worksheet
        startRow = InsertImages(ws, imgNodes, mimeImages, startRow);

        // Write title row if available
        if (!string.IsNullOrWhiteSpace(reportTitle))
        {
            var cell = ws.Cell(startRow, 1);
            cell.Value = reportTitle;
            cell.Style.Font.Bold = true;
            cell.Style.Font.FontSize = 16;
            startRow++;
        }

        // Write description row if available
        if (!string.IsNullOrWhiteSpace(reportDesc))
        {
            var cell = ws.Cell(startRow, 1);
            cell.Value = reportDesc;
            cell.Style.Font.FontSize = 10;
            cell.Style.Font.FontColor = XLColor.Gray;
            startRow++;
        }

        if (startRow > 1)
            startRow++; // blank row between header info and data

        // Process each row
        for (int r = 0; r < rows.Count; r++)
        {
            var tr = rows[r];
            var rowClass = GetClass(tr);
            var cells = tr.SelectNodes("td|th");
            if (cells == null) continue;

            for (int c = 0; c < cells.Count; c++)
            {
                var td = cells[c];
                var xlCell = ws.Cell(startRow + r, c + 1);

                // Extract the text value (strip nested divs/spans/nobrs)
                var rawText = HtmlEntity.DeEntitize(td.InnerText).Trim();

                // Determine the format class from td or child div
                var formatClass = DetectFormatClass(td);

                // Set cell value based on detected format
                SetCellValue(xlCell, rawText, formatClass);

                // Apply alignment
                ApplyAlignment(xlCell, td);

                // Apply row-level styling
                ApplyRowStyle(xlCell, rowClass);
            }
        }

        // Auto-fit columns
        ws.Columns().AdjustToContents();

        // Cap column widths at 50 characters
        foreach (var col in ws.ColumnsUsed())
        {
            if (col.Width > 50)
                col.Width = 50;
        }

        wb.SaveAs(outputPath);
    }

    /// <summary>
    /// Detect which xls-* format class applies to a cell, checking both the td and any child div.
    /// </summary>
    private string? DetectFormatClass(HtmlNode td)
    {
        // Check td's own class
        var tdClass = GetClass(td);
        var match = FindFormatClass(tdClass);
        if (match != null) return match;

        // Check child divs
        var childDivs = td.SelectNodes(".//div[@class]");
        if (childDivs != null)
        {
            foreach (var div in childDivs)
            {
                match = FindFormatClass(GetClass(div));
                if (match != null) return match;
            }
        }

        return null;
    }

    /// <summary>
    /// Find a matching format class key from a CSS class string.
    /// </summary>
    private string? FindFormatClass(string cssClasses)
    {
        if (string.IsNullOrEmpty(cssClasses)) return null;

        // Check longer keys first to avoid partial matches (e.g., "xls-date-ShortDateFormat" before "xls-date")
        foreach (var key in FormatMap.Keys.OrderByDescending(k => k.Length))
        {
            if (cssClasses.Contains(key, StringComparison.OrdinalIgnoreCase))
                return key;
        }
        return null;
    }

    /// <summary>
    /// Set the cell value and number format based on the detected format class.
    /// </summary>
    private void SetCellValue(IXLCell cell, string rawText, string? formatClass)
    {
        if (string.IsNullOrWhiteSpace(rawText) || rawText == "&nbsp;" || rawText == "\u00A0")
        {
            cell.Value = Blank.Value;
            return;
        }

        // Date/time formats: value is an OLE Automation date serial number
        if (formatClass != null && formatClass.StartsWith("xls-date", StringComparison.OrdinalIgnoreCase)
            || formatClass == "xls-time")
        {
            if (double.TryParse(rawText, NumberStyles.Float, CultureInfo.InvariantCulture, out var serial))
            {
                try
                {
                    cell.Value = DateTime.FromOADate(serial);
                    cell.Style.NumberFormat.Format = FormatMap[formatClass!];
                    return;
                }
                catch
                {
                    // Fall through to text
                }
            }
            // If parsing fails, store as text
            cell.Value = rawText;
            cell.Style.NumberFormat.Format = "@";
            return;
        }

        // Percent format
        if (formatClass == "xls-percent")
        {
            if (double.TryParse(rawText.TrimEnd('%'), NumberStyles.Float, CultureInfo.InvariantCulture, out var pct))
            {
                cell.Value = pct / 100.0;
                cell.Style.NumberFormat.Format = FormatMap["xls-percent"];
                return;
            }
        }

        // Text format — force as text
        if (formatClass == "xls-text" || formatClass == "xls-l-text")
        {
            cell.SetValue(rawText);
            cell.Style.NumberFormat.Format = "@";
            return;
        }

        // No format class — try to detect currency or number
        if (TryParseCurrency(rawText, out var amount))
        {
            cell.Value = amount;
            cell.Style.NumberFormat.Format = "$#,##0.00";
            return;
        }

        if (double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands,
                            CultureInfo.InvariantCulture, out var num))
        {
            cell.Value = num;
            return;
        }

        // Default: text
        cell.Value = rawText;
    }

    /// <summary>
    /// Try to parse a US currency string like "$1,234.56" or "-$45.00".
    /// </summary>
    private static bool TryParseCurrency(string text, out decimal amount)
    {
        amount = 0;
        if (!text.Contains('$')) return false;

        var cleaned = text.Replace("$", "").Replace(",", "").Trim();
        return decimal.TryParse(cleaned, NumberStyles.Float | NumberStyles.AllowLeadingSign,
                                CultureInfo.InvariantCulture, out amount);
    }

    /// <summary>
    /// Apply horizontal alignment from the td's align attribute.
    /// </summary>
    private void ApplyAlignment(IXLCell cell, HtmlNode td)
    {
        var align = td.GetAttributeValue("align", "left").ToLowerInvariant();
        cell.Style.Alignment.Horizontal = align switch
        {
            "center" => XLAlignmentHorizontalValues.Center,
            "right"  => XLAlignmentHorizontalValues.Right,
            _        => XLAlignmentHorizontalValues.Left,
        };
        cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
    }

    /// <summary>
    /// Apply row-level visual styling (font color, background, bold) based on CSS class.
    /// </summary>
    private void ApplyRowStyle(IXLCell cell, string rowClass)
    {
        foreach (var kvp in RowStyles)
        {
            if (rowClass.Contains(kvp.Key, StringComparison.OrdinalIgnoreCase))
            {
                var style = kvp.Value;
                cell.Style.Font.FontColor = style.FontColor;
                cell.Style.Fill.BackgroundColor = style.BackgroundColor;
                if (style.Bold)
                    cell.Style.Font.Bold = true;

                // Header gets a border too
                if (kvp.Key == "ReportHeader")
                {
                    cell.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                    cell.Style.Border.BottomBorderColor = XLColor.White;
                }

                // Footer gets a top double border
                if (kvp.Key == "ReportFooter")
                {
                    cell.Style.Border.TopBorder = XLBorderStyleValues.Double;
                }

                break;
            }
        }

        // Set font family and size from CSS
        cell.Style.Font.FontName = "Tahoma";
        if (rowClass.Contains("ReportHeader", StringComparison.OrdinalIgnoreCase))
            cell.Style.Font.FontSize = 9;
        else
            cell.Style.Font.FontSize = 8;
    }

    /// <summary>
    /// Get the class attribute value from an HTML node.
    /// </summary>
    private static string GetClass(HtmlNode node)
    {
        return node.GetAttributeValue("class", "");
    }

    /// <summary>
    /// Sanitize a string for use as an Excel sheet name.
    /// </summary>
    private static string SanitizeSheetName(string name)
    {
        // Excel sheet names cannot contain: \ / ? * [ ] :
        var sanitized = Regex.Replace(name, @"[\\/?*\[\]:]", "_");
        return sanitized.Length > 31 ? sanitized[..31] : sanitized;
    }

    // =====================================================================
    //  MIME image extraction and embedding
    // =====================================================================

    /// <summary>
    /// Parse MIME parts after the closing &lt;/html&gt; tag to extract base64-encoded images.
    /// Returns a dictionary mapping Content-ID → decoded image bytes.
    /// </summary>
    private static Dictionary<string, byte[]> ExtractMimeImages(string rawContent)
    {
        var images = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);

        // Find content after </html> — that's where MIME parts live
        var htmlEnd = rawContent.IndexOf("</html>", StringComparison.OrdinalIgnoreCase);
        if (htmlEnd < 0) return images;

        var mimeSection = rawContent[(htmlEnd + "</html>".Length)..];

        // Match MIME parts: Content-ID line, Content-Transfer-Encoding line, blank line, then base64 data
        var mimePartPattern = new Regex(
            @"Content-ID:\s*(?<cid>\S+)\s*\r?\nContent-Transfer-Encoding:\s*BASE64\s*\r?\n\s*\r?\n(?<data>[A-Za-z0-9+/=\s]+?)(?=\r?\nContent-ID:|\s*$)",
            RegexOptions.IgnoreCase | RegexOptions.Singleline);

        foreach (Match m in mimePartPattern.Matches(mimeSection))
        {
            var cid = m.Groups["cid"].Value.Trim();
            var base64 = m.Groups["data"].Value.Trim();

            // Remove any whitespace/newlines from the base64 data
            base64 = Regex.Replace(base64, @"\s+", "");

            try
            {
                var bytes = System.Convert.FromBase64String(base64);
                images[cid] = bytes;
            }
            catch (FormatException)
            {
                // Skip malformed base64 data
            }
        }

        return images;
    }

    /// <summary>
    /// Represents an image reference found in the HTML via a cid: src attribute.
    /// </summary>
    private record CidImageRef(string ContentId, int Width, int Height);

    /// <summary>
    /// Find all &lt;img&gt; tags in the document that reference embedded images via cid: URLs.
    /// </summary>
    private static List<CidImageRef> FindCidImages(HtmlDocument doc)
    {
        var result = new List<CidImageRef>();

        var imgNodes = doc.DocumentNode.SelectNodes("//img[contains(@src,'cid:')]");
        if (imgNodes == null) return result;

        foreach (var img in imgNodes)
        {
            var src = img.GetAttributeValue("src", "");
            if (!src.StartsWith("cid:", StringComparison.OrdinalIgnoreCase)) continue;

            var cid = src["cid:".Length..].Trim();

            // Try to get dimensions from width/height attributes first
            var width = img.GetAttributeValue("width", 0);
            var height = img.GetAttributeValue("height", 0);

            // If no attributes, try to parse from inline style (e.g., "width: 1000px; height: 300px")
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

    /// <summary>
    /// Insert extracted MIME images into the worksheet, returning the next available row.
    /// </summary>
    private static int InsertImages(IXLWorksheet ws, List<CidImageRef> imgRefs,
                                     Dictionary<string, byte[]> mimeImages, int startRow)
    {
        if (imgRefs.Count == 0 || mimeImages.Count == 0)
            return startRow;

        int imageIndex = 0;
        foreach (var imgRef in imgRefs)
        {
            if (!mimeImages.TryGetValue(imgRef.ContentId, out var imageBytes))
                continue;

            imageIndex++;
            using var ms = new MemoryStream(imageBytes);

            var picture = ws.AddPicture(ms, $"Image_{imageIndex}")
                .MoveTo(ws.Cell(startRow, 1));

            // Scale large images down to fit reasonably in the spreadsheet
            // Excel column width ~7.5px per unit, typical visible width ~750px for 10 columns
            if (imgRef.Width > 0 && imgRef.Height > 0)
            {
                // Use original dimensions but cap width at 750px
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
                // No dimensions known — scale to 50% as a reasonable default
                picture.Scale(0.5);
            }

            // Estimate how many rows the image occupies (default row height ~15px)
            var imgHeight = imgRef.Height > 0 ? imgRef.Height : 200;
            // Apply same scaling as the picture
            const int maxW = 750;
            if (imgRef.Width > maxW && imgRef.Width > 0)
                imgHeight = (int)(imgHeight * ((double)maxW / imgRef.Width));
            int rowsForImage = Math.Max(1, (int)Math.Ceiling(imgHeight / 15.0));

            startRow += rowsForImage + 1; // +1 for spacing
        }

        return startRow;
    }
}

/// <summary>
/// Simple record to hold row styling info.
/// </summary>
record RowStyle(XLColor FontColor, XLColor BackgroundColor, bool Bold);
