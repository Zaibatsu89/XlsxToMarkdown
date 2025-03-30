using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using System.Text;

namespace XlsxToMarkdown
{
    public class Program
    {
        public static async Task Main(string[] args)
        {
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: XlsxToMarkdown <input.xlsx> <output.md>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];

            try
            {
                await ConvertXlsxToMarkdownAsync(inputPath, outputPath);
                Console.WriteLine($"Conversion complete: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }

        private static async Task ConvertXlsxToMarkdownAsync(string inputPath, string outputPath)
        {
            if (!File.Exists(inputPath))
            {
                throw new FileNotFoundException("Input XLSX file not found.", inputPath);
            }

            StringBuilder markdownBuilder = new();

            // Add metadata section
            markdownBuilder.AppendLine("# Excel Document Conversion");
            markdownBuilder.AppendLine($"*Source: {Path.GetFileName(inputPath)}*");
            markdownBuilder.AppendLine($"*Converted: {DateTime.Now:yyyy-MM-dd HH:mm:ss}*");
            markdownBuilder.AppendLine();

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(inputPath, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart!;
                SharedStringTablePart? sharedStringPart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                SharedStringTable? sharedStringTable = sharedStringPart?.SharedStringTable;

                // Process all worksheets
                IEnumerable<Sheet> sheets = workbookPart.Workbook.Descendants<Sheet>();

                foreach (Sheet sheet in sheets)
                {
                    string? relationshipId = sheet.Id?.Value;
                    if (relationshipId == null)
                    {
                        continue;
                    }

                    WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(relationshipId);

                    // Add sheet header
                    markdownBuilder.AppendLine($"## {sheet.Name}");
                    markdownBuilder.AppendLine();

                    // Process content
                    ProcessWorksheet(worksheetPart, sharedStringTable, markdownBuilder);
                    markdownBuilder.AppendLine();
                }

                // Extract document properties if available
                ExtractMetadata(spreadsheetDocument, markdownBuilder);
            }

            await File.WriteAllTextAsync(outputPath, markdownBuilder.ToString());
        }

        private static void ProcessWorksheet(WorksheetPart worksheetPart, SharedStringTable? sharedStringTable, StringBuilder markdownBuilder)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>()!;

            if (!sheetData.HasChildren)
            {
                markdownBuilder.AppendLine("*Empty worksheet*");
                return;
            }

            // Determine table dimensions
            var rows = sheetData.Elements<Row>().ToList();
            if (!rows.Any())
                return;

            // Calculate max column count
            int maxColumnCount = rows.Max(r => r.Elements<Cell>().Count());

            // First row as header
            Row? headerRow = rows.FirstOrDefault();
            if (headerRow != null)
            {
                markdownBuilder.Append("| ");
                for (int i = 0; i < maxColumnCount; i++)
                {
                    Cell? cell = GetCellByColumnIndex(headerRow, i);
                    string cellValue = GetCellValue(cell, sharedStringTable) ?? "";
                    markdownBuilder.Append(EscapeMarkdownChars(cellValue));
                    markdownBuilder.Append(" | ");
                }
                markdownBuilder.AppendLine();

                // Add separator row
                markdownBuilder.Append("| ");
                for (int i = 0; i < maxColumnCount; i++)
                {
                    markdownBuilder.Append("--- | ");
                }
                markdownBuilder.AppendLine();
            }

            // Process data rows (skip header)
            foreach (var row in rows.Skip(1))
            {
                markdownBuilder.Append("| ");
                for (int i = 0; i < maxColumnCount; i++)
                {
                    Cell? cell = GetCellByColumnIndex(row, i);
                    string cellValue = GetCellValue(cell, sharedStringTable) ?? "";
                    markdownBuilder.Append(EscapeMarkdownChars(cellValue));
                    markdownBuilder.Append(" | ");
                }
                markdownBuilder.AppendLine();
            }
        }

        private static Cell? GetCellByColumnIndex(Row row, int columnIndex)
        {
            // Convert zero-based index to Excel column reference (A, B, C, ...)
            string columnReference = GetColumnReference(columnIndex);

            // Try to find the cell
            return row.Elements<Cell>()
                .FirstOrDefault(c => string.Equals(
                    GetColumnFromCellReference(c.CellReference?.Value),
                    columnReference,
                    StringComparison.OrdinalIgnoreCase));
        }

        private static string GetColumnReference(int columnIndex)
        {
            int dividend = columnIndex + 1; // 1-based
            string columnName = string.Empty;

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }

        private static string? GetColumnFromCellReference(string? cellReference)
        {
            if (string.IsNullOrEmpty(cellReference))
                return null;

            // Extract column letters from cell reference (e.g., "A1" -> "A")
            return new string(cellReference.TakeWhile(char.IsLetter).ToArray());
        }

        private static string GetCellValue(Cell? cell, SharedStringTable? sharedStringTable)
        {
            if (cell == null)
                return string.Empty;

            // If no DataType is specified, return direct value
            if (cell.DataType == null)
                return cell.CellValue?.Text ?? string.Empty;

            // Apply appropriate conversion based on DataType
            return ConvertCellValue(cell, sharedStringTable) ?? string.Empty;
        }

        private static string? ConvertCellValue(Cell? cell, SharedStringTable? sharedStringTable)
        {
            if (cell?.DataType == null || cell.CellValue == null)
                return null;

            // Using object pattern with when clause for enum comparison
            return cell.DataType switch
            {
                object o when o.Equals(CellValues.Boolean) => cell.CellValue.Text == "1" ? "True" : "False",
                object o when o.Equals(CellValues.Date) => DateTime.FromOADate(double.Parse(cell.CellValue.Text)).ToString("yyyy-MM-dd"),
                object o when o.Equals(CellValues.SharedString) => GetSharedStringItemById(sharedStringTable, int.Parse(cell.CellValue.Text)),
                object o when o.Equals(CellValues.Number) => cell.CellValue.Text,
                object o when o.Equals(CellValues.String) => cell.CellValue.Text,
                _ => cell.CellValue.Text
            };
        }

        private static string? GetSharedStringItemById(SharedStringTable? sharedStringTable, int id)
        {
            if (sharedStringTable == null || id < 0 || id >= sharedStringTable.Count())
                return null;

            SharedStringItem? item = sharedStringTable.Elements<SharedStringItem>().ElementAtOrDefault(id);
            return item?.Text?.Text ?? item?.InnerText ?? string.Empty;
        }

        private static string EscapeMarkdownChars(string text)
        {
            // Escape pipe characters and other markdown special characters
            return text
                .Replace("|", "\\|")
                .Replace("\n", " ") // Replace newlines with spaces for better table formatting
                .Trim();
        }

        private static void ExtractMetadata(SpreadsheetDocument doc, StringBuilder markdownBuilder)
        {
            var coreProps = doc.PackageProperties;
            if (coreProps != null)
            {
                markdownBuilder.AppendLine("## Document Metadata");

                if (!string.IsNullOrEmpty(coreProps.Title))
                {
                    markdownBuilder.AppendLine($"**Title**: {coreProps.Title}");
                }

                if (!string.IsNullOrEmpty(coreProps.Subject))
                {
                    markdownBuilder.AppendLine($"**Subject**: {coreProps.Subject}");
                }

                if (!string.IsNullOrEmpty(coreProps.Creator))
                {
                    markdownBuilder.AppendLine($"**Author**: {coreProps.Creator}");
                }

                if (coreProps.Created.HasValue)
                {
                    markdownBuilder.AppendLine($"**Created**: {coreProps.Created.Value:yyyy-MM-dd HH:mm:ss}");
                }

                markdownBuilder.AppendLine();
            }
        }
    }
}