using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using iTextSharp.text;
using iTextSharp.text.pdf;

public class MicrosoftDataAnalysisExample
{
    public static void Main(string[] args)
    {
        try
        {
            Console.WriteLine("Microsoft.Data.Analysis Implementation");
            Console.WriteLine(new string('=', 60));
            
            string excelFilePath = args.Length > 0 ? args[0] : "test.xlsx";
            Console.WriteLine($"Reading Excel file: {excelFilePath}");

            // Read the Excel file using our ExcelReader, then convert to Microsoft DataFrame
            var customReader = new ExcelReader();
            var customDf = customReader.ReadDataFrame(excelFilePath);

            Console.WriteLine($"Custom DataFrame read: {customDf.RowCount} rows × {customDf.ColumnCount} columns");

            // Convert our custom DataFrame to Microsoft.Data.Analysis DataFrame
            var microsoftDf = ConvertToMicrosoftDataFrame(customDf);

            Console.WriteLine("Microsoft DataFrame created successfully!");
            Console.WriteLine($"Rows: {microsoftDf.Rows.Count}");
            Console.WriteLine($"Columns: {microsoftDf.Columns.Count}");
            
            Console.WriteLine("\nMicrosoft DataFrame content:");
            Console.WriteLine(microsoftDf);

            // Test statistics on the same data
            Console.WriteLine("\nStatistics (Microsoft.Data.Analysis):");
            for (int i = 0; i < microsoftDf.Columns.Count; i++)
            {
                var column = microsoftDf.Columns[i];
                if (column is Microsoft.Data.Analysis.PrimitiveDataFrameColumn<double> doubleCol)
                {
                    var mean = doubleCol.Mean();
                    var min = doubleCol.Min();
                    var max = doubleCol.Max();
                    Console.WriteLine($"  {column.Name}: Count={doubleCol.Length}, Mean={mean:F2}, Min={min}, Max={max}");
                }
                else
                {
                    Console.WriteLine($"  {column.Name}: Non-numeric column (Count={column.Length})");
                }
            }

            // Test filtering on the same data
            if (microsoftDf.Columns.Count > 0 && microsoftDf.Columns[0] is Microsoft.Data.Analysis.PrimitiveDataFrameColumn<double> firstCol)
            {
                Console.WriteLine("\nFiltering (Microsoft.Data.Analysis):");
                var mean = firstCol.Mean();
                var filter = firstCol.ElementwiseGreaterThan(mean);
                var filtered = microsoftDf.Filter(filter);
                Console.WriteLine($"Filtered ('{firstCol.Name}' > {mean:F2}):");
                Console.WriteLine(filtered);
            }

            // Try sorting
            if (microsoftDf.Columns.Count > 0 && microsoftDf.Columns[0] is Microsoft.Data.Analysis.PrimitiveDataFrameColumn<double> sortCol)
            {
                Console.WriteLine("\nSorting (Microsoft.Data.Analysis):");
                var sorted = microsoftDf.OrderBy(sortCol.Name);
                Console.WriteLine($"Sorted by '{sortCol.Name}' (ascending):");
                Console.WriteLine(sorted);
            }

            // Export functionality - Corporate-grade exports
            Console.WriteLine("\nExport Features:");
            
            // Excel Export
            var excelPath = Path.ChangeExtension(excelFilePath, ".microsoft.xlsx");
            ExportToExcel(microsoftDf, excelPath, "Microsoft DataFrame Analysis");
            Console.WriteLine($"✅ Exported to Excel: {excelPath}");
            
            // CSV Export
            var csvPath = Path.ChangeExtension(excelFilePath, ".microsoft.csv");
            ExportToCsv(microsoftDf, csvPath);
            Console.WriteLine($"✅ Exported to CSV: {csvPath}");
            
            // PDF Export
            var pdfPath = Path.ChangeExtension(excelFilePath, ".microsoft.pdf");
            ExportToPdf(microsoftDf, pdfPath, "Corporate Data Analysis Report");
            Console.WriteLine($"✅ Exported to PDF: {pdfPath}");

            Console.WriteLine("\nMicrosoft.Data.Analysis implementation with complete exports completed successfully!");

        }
        catch (Exception ex)
        {
            Console.WriteLine($"Microsoft.Data.Analysis Error: {ex.Message}");
            Console.WriteLine($"Stack trace: {ex.StackTrace}");
        }
    }

    /// <summary>
    /// Export Microsoft DataFrame to Excel format
    /// </summary>
    private static void ExportToExcel(Microsoft.Data.Analysis.DataFrame df, string filePath, string worksheetName = "Data")
    {
        using var package = new ExcelPackage();
        var worksheet = package.Workbook.Worksheets.Add(worksheetName);

        // Write headers
        for (int col = 0; col < df.Columns.Count; col++)
        {
            worksheet.Cells[1, col + 1].Value = df.Columns[col].Name;
            worksheet.Cells[1, col + 1].Style.Font.Bold = true;
        }

        // Write data
        for (long row = 0; row < df.Rows.Count; row++)
        {
            for (int col = 0; col < df.Columns.Count; col++)
            {
                var value = df.Columns[col][row];
                worksheet.Cells[(int)(row + 2), col + 1].Value = value;
            }
        }

        // Auto-fit columns
        worksheet.Cells.AutoFitColumns();

        // Save file
        var file = new FileInfo(filePath);
        package.SaveAs(file);
    }

    /// <summary>
    /// Export Microsoft DataFrame to CSV format
    /// </summary>
    private static void ExportToCsv(Microsoft.Data.Analysis.DataFrame df, string filePath, string delimiter = ",")
    {
        using var writer = new StreamWriter(filePath, false, Encoding.UTF8);
        
        // Write headers
        var headers = new List<string>();
        for (int i = 0; i < df.Columns.Count; i++)
        {
            headers.Add(df.Columns[i].Name);
        }
        writer.WriteLine(string.Join(delimiter, headers.Select(h => EscapeCsvField(h, delimiter))));
        
        // Write data rows
        for (long row = 0; row < df.Rows.Count; row++)
        {
            var values = new List<string>();
            for (int col = 0; col < df.Columns.Count; col++)
            {
                var value = df.Columns[col][row]?.ToString() ?? "";
                values.Add(EscapeCsvField(value, delimiter));
            }
            writer.WriteLine(string.Join(delimiter, values));
        }
    }

    /// <summary>
    /// Export Microsoft DataFrame to PDF format
    /// </summary>
    private static void ExportToPdf(Microsoft.Data.Analysis.DataFrame df, string filePath, string title = "DataFrame Report")
    {
        using var stream = new FileStream(filePath, FileMode.Create);
        var document = new Document(PageSize.A4.Rotate(), 10f, 10f, 10f, 10f);
        var writer = PdfWriter.GetInstance(document, stream);
        
        document.Open();
        
        // Add title
        var titleFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 16);
        var titleParagraph = new Paragraph(title, titleFont)
        {
            Alignment = Element.ALIGN_CENTER,
            SpacingAfter = 20f
        };
        document.Add(titleParagraph);
        
        // Create table
        var table = new PdfPTable(df.Columns.Count)
        {
            WidthPercentage = 100
        };
        
        // Add headers
        var headerFont = FontFactory.GetFont(FontFactory.HELVETICA_BOLD, 10);
        for (int i = 0; i < df.Columns.Count; i++)
        {
            var cell = new PdfPCell(new Phrase(df.Columns[i].Name, headerFont))
            {
                BackgroundColor = new BaseColor(230, 230, 230),
                HorizontalAlignment = Element.ALIGN_CENTER,
                Padding = 5f
            };
            table.AddCell(cell);
        }
        
        // Add data rows
        var dataFont = FontFactory.GetFont(FontFactory.HELVETICA, 9);
        for (long row = 0; row < df.Rows.Count; row++)
        {
            for (int col = 0; col < df.Columns.Count; col++)
            {
                var value = df.Columns[col][row]?.ToString() ?? "";
                var isNumeric = df.Columns[col] is Microsoft.Data.Analysis.PrimitiveDataFrameColumn<double> ||
                               df.Columns[col] is Microsoft.Data.Analysis.PrimitiveDataFrameColumn<int> ||
                               df.Columns[col] is Microsoft.Data.Analysis.PrimitiveDataFrameColumn<float>;
                
                var cell = new PdfPCell(new Phrase(value, dataFont))
                {
                    Padding = 3f,
                    HorizontalAlignment = isNumeric ? Element.ALIGN_RIGHT : Element.ALIGN_LEFT
                };
                table.AddCell(cell);
            }
        }
        
        document.Add(table);
        
        // Add footer with metadata
        var footerFont = FontFactory.GetFont(FontFactory.HELVETICA, 8);
        var footer = new Paragraph($"\nGenerated on: {DateTime.Now:yyyy-MM-dd HH:mm:ss} | Rows: {df.Rows.Count} | Columns: {df.Columns.Count}", footerFont)
        {
            Alignment = Element.ALIGN_CENTER,
            SpacingBefore = 20f
        };
        document.Add(footer);
        
        document.Close();
    }

    /// <summary>
    /// Helper method to escape CSV fields that contain delimiters, quotes, or newlines
    /// </summary>
    private static string EscapeCsvField(string field, string delimiter)
    {
        if (field.Contains(delimiter) || field.Contains("\"") || field.Contains("\n") || field.Contains("\r"))
        {
            return "\"" + field.Replace("\"", "\"\"") + "\"";
        }
        return field;
    }

    private static Microsoft.Data.Analysis.DataFrame ConvertToMicrosoftDataFrame(DataFrame customDf)
    {
        var columns = new List<Microsoft.Data.Analysis.DataFrameColumn>();

        foreach (var header in customDf.Headers)
        {
            // Get all values for this column
            var columnValues = customDf.GetColumn(header);
            
            // Check if all non-null values are numeric
            var nonNullValues = columnValues.Where(v => v != null).ToList();
            bool isNumeric = nonNullValues.All(v => v is double || v is int || v is long || v is decimal);

            if (isNumeric && nonNullValues.Count > 0)
            {
                // Create numeric column
                var doubleValues = columnValues.Select(v => v == null ? 0.0 : Convert.ToDouble(v)).ToList();
                columns.Add(new Microsoft.Data.Analysis.PrimitiveDataFrameColumn<double>(header, doubleValues));
            }
            else
            {
                // Create string column
                var stringValues = columnValues.Select(v => v?.ToString() ?? "").ToList();
                columns.Add(new Microsoft.Data.Analysis.StringDataFrameColumn(header, stringValues));
            }
        }

        return new Microsoft.Data.Analysis.DataFrame(columns);
    }
}

// Minimal DataFrame and ExcelReader classes for this implementation
public class DataFrame
{
    public List<string> Headers { get; set; } = new List<string>();
    public List<Dictionary<string, object>> Rows { get; set; } = new List<Dictionary<string, object>>();
    public int RowCount => Rows.Count;
    public int ColumnCount => Headers.Count;
    
    public void AddRow(Dictionary<string, object> row) => Rows.Add(row);
    
    public List<object?> GetColumn(string columnName)
    {
        return Rows.Select(row => row.ContainsKey(columnName) ? row[columnName] : null).ToList();
    }
}

public class ExcelReader
{
    static ExcelReader()
    {
        ExcelPackage.License.SetNonCommercialPersonal("Personal Use");
    }

    public DataFrame ReadDataFrame(string filePath)
    {
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"Excel file not found: {filePath}");

        var dataFrame = new DataFrame();

        using var package = new ExcelPackage(new FileInfo(filePath));
        var worksheet = package.Workbook.Worksheets[0];
        
        if (worksheet.Dimension == null) return dataFrame;

        int rowCount = worksheet.Dimension.Rows;
        int colCount = worksheet.Dimension.Columns;

        // Read headers
        for (int col = 1; col <= colCount; col++)
        {
            var headerValue = worksheet.Cells[1, col].Value?.ToString()?.Trim() ?? $"Column{col}";
            dataFrame.Headers.Add(headerValue);
        }

        // Read data rows
        for (int row = 2; row <= rowCount; row++)
        {
            var rowData = new Dictionary<string, object>();
            bool hasData = false;

            for (int col = 1; col <= colCount; col++)
            {
                var cellValue = worksheet.Cells[row, col].Value;
                var headerName = dataFrame.Headers[col - 1];
                
                rowData[headerName] = cellValue ?? "";
                if (cellValue != null) hasData = true;
            }

            if (hasData) dataFrame.AddRow(rowData);
        }

        return dataFrame;
    }
}