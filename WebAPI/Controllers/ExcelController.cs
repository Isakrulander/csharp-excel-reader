using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using Microsoft.Data.Analysis;
using System.Globalization;
using System.Text;
using PdfSharp.Pdf;
using PdfSharp.Drawing;

namespace WebAPIClean.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ExcelController : ControllerBase
    {
        private readonly ILogger<ExcelController> _logger;

        public ExcelController(ILogger<ExcelController> logger)
        {
            _logger = logger;
        }

        [HttpGet("health")]
        public IActionResult Health()
        {
            return Ok(new { status = "healthy", message = "Excel API is running" });
        }

        [HttpPost("export/csv")]
        public async Task<IActionResult> ExportToCsv(IFormFile file)
        {
            try
            {
                var dataFrame = await ProcessExcelFile(file);
                var csvContent = ExportDataFrameToCsv(dataFrame);
                
                var fileName = Path.GetFileNameWithoutExtension(file.FileName) + ".csv";
                return File(Encoding.UTF8.GetBytes(csvContent), "text/csv", fileName);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error exporting to CSV");
                return StatusCode(500, new { error = "Error exporting to CSV", details = ex.Message });
            }
        }

        [HttpPost("export/excel")]
        public async Task<IActionResult> ExportToExcel(IFormFile file)
        {
            try
            {
                var dataFrame = await ProcessExcelFile(file);
                var excelContent = ExportDataFrameToExcel(dataFrame);
                
                var fileName = Path.GetFileNameWithoutExtension(file.FileName) + "_processed.xlsx";
                return File(excelContent, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error exporting to Excel");
                return StatusCode(500, new { error = "Error exporting to Excel", details = ex.Message });
            }
        }

        [HttpPost("export/pdf")]
        public async Task<IActionResult> ExportToPdf(IFormFile file)
        {
            try
            {
                var dataFrame = await ProcessExcelFile(file);
                var pdfContent = ExportDataFrameToPdf(dataFrame, file.FileName);
                
                var fileName = Path.GetFileNameWithoutExtension(file.FileName) + ".pdf";
                return File(pdfContent, "application/pdf", fileName);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error exporting to PDF");
                return StatusCode(500, new { error = "Error exporting to PDF", details = ex.Message });
            }
        }

        [HttpPost("upload")]
        public async Task<IActionResult> UploadExcel(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest(new { error = "No file uploaded" });
            }

            if (!file.FileName.EndsWith(".xlsx") && !file.FileName.EndsWith(".xls"))
            {
                return BadRequest(new { error = "Please upload an Excel file (.xlsx or .xls)" });
            }

            try
            {
                using var stream = new MemoryStream();
                await file.CopyToAsync(stream);
                
                using var package = new ExcelPackage(stream);
                var worksheet = package.Workbook.Worksheets[0];
                
                // Create DataFrame for advanced data analysis
                var dataFrame = CreateDataFrameFromWorksheet(worksheet);
                
                var result = new
                {
                    fileName = file.FileName,
                    worksheetName = worksheet.Name,
                    rowCount = (int)dataFrame.Rows.Count,
                    columnCount = dataFrame.Columns.Count,
                    
                    // Corporate-grade data analysis
                    columns = GetColumnInfo(dataFrame),
                    statistics = GetDataFrameStatistics(dataFrame),
                    dataTypes = GetDataTypes(dataFrame),
                    
                    // Data preview (first 10 rows)
                    preview = GetDataPreview(dataFrame, 10),
                    
                    // Advanced analytics
                    summary = GetDataSummary(dataFrame)
                };

                _logger.LogInformation($"Successfully analyzed Excel file: {file.FileName} with {dataFrame.Rows.Count} rows and {dataFrame.Columns.Count} columns");
                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error processing Excel file: {file.FileName}");
                return StatusCode(500, new { error = "Error processing Excel file", details = ex.Message });
            }
        }

        private DataFrame CreateDataFrameFromWorksheet(ExcelWorksheet worksheet)
        {
            if (worksheet.Dimension == null)
                return new DataFrame();

            var columns = new List<DataFrameColumn>();
            var headers = new List<string>();

            // Get headers from first row
            for (int col = 1; col <= worksheet.Dimension.Columns; col++)
            {
                headers.Add(worksheet.Cells[1, col].Text ?? $"Column{col}");
            }

            // Create columns and determine data types
            for (int colIndex = 0; colIndex < headers.Count; colIndex++)
            {
                var columnData = new List<object>();
                bool isNumeric = true;
                bool isDateTime = true;

                // Analyze data types by checking first few non-empty values
                for (int row = 2; row <= Math.Min(worksheet.Dimension.Rows, 11); row++)
                {
                    var cellValue = worksheet.Cells[row, colIndex + 1].Text;
                    if (!string.IsNullOrEmpty(cellValue))
                    {
                        if (!double.TryParse(cellValue, out _))
                            isNumeric = false;
                        if (!DateTime.TryParse(cellValue, out _))
                            isDateTime = false;
                    }
                }

                // Collect all data for this column
                for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                {
                    var cellValue = worksheet.Cells[row, colIndex + 1].Text ?? "";
                    
                    if (isNumeric && double.TryParse(cellValue, out double numValue))
                    {
                        columnData.Add(numValue);
                    }
                    else if (isDateTime && DateTime.TryParse(cellValue, out DateTime dateValue))
                    {
                        columnData.Add(dateValue);
                    }
                    else
                    {
                        columnData.Add(cellValue);
                    }
                }

                // Create appropriate column type
                if (isNumeric)
                {
                    var doubleValues = columnData.Select(x => x is double d ? d : double.TryParse(x?.ToString(), out var parsed) ? parsed : (double?)null);
                    columns.Add(new PrimitiveDataFrameColumn<double>(headers[colIndex], doubleValues));
                }
                else if (isDateTime)
                {
                    var dateValues = columnData.Select(x => x is DateTime dt ? dt : DateTime.TryParse(x?.ToString(), out var parsed) ? parsed : (DateTime?)null);
                    columns.Add(new PrimitiveDataFrameColumn<DateTime>(headers[colIndex], dateValues));
                }
                else
                {
                    var stringValues = columnData.Select(x => x?.ToString() ?? "");
                    columns.Add(new StringDataFrameColumn(headers[colIndex], stringValues));
                }
            }

            return new DataFrame(columns);
        }

        private object GetColumnInfo(DataFrame df)
        {
            return df.Columns.Select(col => new
            {
                name = col.Name,
                type = col.DataType.Name,
                nullCount = col.NullCount,
                length = col.Length
            }).ToList();
        }

        private object GetDataFrameStatistics(DataFrame df)
        {
            var stats = new Dictionary<string, object>();
            
            foreach (var column in df.Columns)
            {
                if (column is PrimitiveDataFrameColumn<double> numCol)
                {
                    var mean = numCol.Mean();
                    var min = numCol.Min();
                    var max = numCol.Max();
                    
                    stats[column.Name] = new
                    {
                        mean = mean,
                        min = min,
                        max = max,
                        count = numCol.Length - numCol.NullCount,
                        type = "numeric"
                    };
                }
                else if (column is StringDataFrameColumn strCol)
                {
                    var valueCounts = strCol.ValueCounts();
                    stats[column.Name] = new
                    {
                        uniqueCount = valueCounts.Rows.Count,
                        nullCount = strCol.NullCount,
                        totalCount = strCol.Length,
                        type = "text"
                    };
                }
            }
            
            return stats;
        }

        private object GetDataTypes(DataFrame df)
        {
            return df.Columns.ToDictionary(col => col.Name, col => col.DataType.Name);
        }

        private List<Dictionary<string, object>> GetDataPreview(DataFrame df, int rows)
        {
            var preview = new List<Dictionary<string, object>>();
            var rowCount = Math.Min(rows, (int)df.Rows.Count);
            
            for (int i = 0; i < rowCount; i++)
            {
                var row = new Dictionary<string, object>();
                foreach (var column in df.Columns)
                {
                    row[column.Name] = column[i] ?? "";
                }
                preview.Add(row);
            }
            
            return preview;
        }

        private object GetDataSummary(DataFrame df)
        {
            return new
            {
                totalRows = df.Rows.Count,
                totalColumns = df.Columns.Count,
                numericColumns = df.Columns.Count(c => c is PrimitiveDataFrameColumn<double>),
                textColumns = df.Columns.Count(c => c is StringDataFrameColumn),
                dateColumns = df.Columns.Count(c => c is PrimitiveDataFrameColumn<DateTime>),
                memoryUsage = $"{df.Columns.Sum(c => c.Length * 8)} bytes (estimated)"
            };
        }

        private async Task<DataFrame> ProcessExcelFile(IFormFile file)
        {
            using var stream = new MemoryStream();
            await file.CopyToAsync(stream);
            
            using var package = new ExcelPackage(stream);
            var worksheet = package.Workbook.Worksheets[0];
            
            return CreateDataFrameFromWorksheet(worksheet);
        }

        private string ExportDataFrameToCsv(DataFrame df)
        {
            var csv = new StringBuilder();
            
            // Add headers
            csv.AppendLine(string.Join(",", df.Columns.Select(c => c.Name)));
            
            // Add data rows
            for (int i = 0; i < df.Rows.Count; i++)
            {
                var values = df.Columns.Select(col => {
                    var value = col[i]?.ToString() ?? "";
                    // Escape values containing commas or quotes
                    if (value.Contains(",") || value.Contains("\""))
                    {
                        value = "\"" + value.Replace("\"", "\"\"") + "\"";
                    }
                    return value;
                });
                csv.AppendLine(string.Join(",", values));
            }
            
            return csv.ToString();
        }

        private byte[] ExportDataFrameToExcel(DataFrame df)
        {
            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("ProcessedData");
            
            // Add headers
            for (int col = 0; col < df.Columns.Count; col++)
            {
                worksheet.Cells[1, col + 1].Value = df.Columns[col].Name;
                worksheet.Cells[1, col + 1].Style.Font.Bold = true;
            }
            
            // Add data
            for (int row = 0; row < df.Rows.Count; row++)
            {
                for (int col = 0; col < df.Columns.Count; col++)
                {
                    worksheet.Cells[row + 2, col + 1].Value = df.Columns[col][row];
                }
            }
            
            // Auto-fit columns
            worksheet.Cells.AutoFitColumns();
            
            return package.GetAsByteArray();
        }

        private byte[] ExportDataFrameToPdf(DataFrame df, string fileName)
        {
            try
            {
                // Create a new PDF document
                var document = new PdfDocument();
                var page = document.AddPage();
                var gfx = XGraphics.FromPdfPage(page);
                
                // Define fonts
                var titleFont = new XFont("Arial", 16, XFontStyleEx.Bold);
                var headerFont = new XFont("Arial", 12, XFontStyleEx.Bold);
                var textFont = new XFont("Arial", 10, XFontStyleEx.Regular);
                var dataFont = new XFont("Arial", 9, XFontStyleEx.Regular);
                
                double yPosition = 40;
                double leftMargin = 40;
                double rightMargin = page.Width - 40;
                
                // Title
                gfx.DrawString("DATA ANALYSIS REPORT", titleFont, XBrushes.Black, leftMargin, yPosition);
                yPosition += 30;
                
                // File information
                gfx.DrawString($"File: {fileName}", textFont, XBrushes.Black, leftMargin, yPosition);
                yPosition += 15;
                gfx.DrawString($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}", textFont, XBrushes.Gray, leftMargin, yPosition);
                yPosition += 15;
                gfx.DrawString($"Rows: {df.Rows.Count:N0}, Columns: {df.Columns.Count:N0}", textFont, XBrushes.Black, leftMargin, yPosition);
                yPosition += 25;
                
                // Column Information
                gfx.DrawString("COLUMN INFORMATION", headerFont, XBrushes.Black, leftMargin, yPosition);
                yPosition += 20;
                
                foreach (var column in df.Columns)
                {
                    if (yPosition > page.Height - 100) break; // Stop if near bottom
                    
                    var info = $"• {column.Name} ({column.DataType.Name}) - Nulls: {column.NullCount:N0}";
                    gfx.DrawString(info, dataFont, XBrushes.Black, leftMargin + 10, yPosition);
                    yPosition += 12;
                }
                
                yPosition += 10;
                
                // Statistics for numeric columns
                var numericColumns = df.Columns.Where(c => c is PrimitiveDataFrameColumn<double>).ToList();
                if (numericColumns.Any() && yPosition < page.Height - 150)
                {
                    gfx.DrawString("NUMERIC STATISTICS", headerFont, XBrushes.Black, leftMargin, yPosition);
                    yPosition += 20;
                    
                    foreach (var col in numericColumns.Cast<PrimitiveDataFrameColumn<double>>())
                    {
                        if (yPosition > page.Height - 100) break;
                        
                        var mean = col.Mean();
                        var min = col.Min();
                        var max = col.Max();
                        var stats = $"• {col.Name}: Mean={mean:F2}, Min={min:F2}, Max={max:F2}";
                        gfx.DrawString(stats, dataFont, XBrushes.Black, leftMargin + 10, yPosition);
                        yPosition += 12;
                    }
                    yPosition += 10;
                }
                
                // Data Preview
                if (yPosition < page.Height - 100)
                {
                    gfx.DrawString("DATA PREVIEW", headerFont, XBrushes.Black, leftMargin, yPosition);
                    yPosition += 20;
                    
                    // Headers
                    double colWidth = (rightMargin - leftMargin) / df.Columns.Count;
                    double xPos = leftMargin;
                    
                    foreach (var column in df.Columns)
                    {
                        var colName = column.Name.Length > 10 ? column.Name.Substring(0, 7) + "..." : column.Name;
                        gfx.DrawString(colName, dataFont, XBrushes.Black, xPos, yPosition);
                        xPos += colWidth;
                    }
                    yPosition += 15;
                    
                    // Data rows (first few that fit)
                    var maxRows = Math.Min((int)df.Rows.Count, (int)((page.Height - yPosition - 50) / 12));
                    
                    for (int row = 0; row < maxRows; row++)
                    {
                        xPos = leftMargin;
                        foreach (var column in df.Columns)
                        {
                            var value = column[row]?.ToString() ?? "";
                            if (value.Length > 10) value = value.Substring(0, 7) + "...";
                            gfx.DrawString(value, dataFont, XBrushes.Black, xPos, yPosition);
                            xPos += colWidth;
                        }
                        yPosition += 12;
                    }
                    
                    if (df.Rows.Count > maxRows)
                    {
                        yPosition += 10;
                        gfx.DrawString($"... showing first {maxRows} of {df.Rows.Count:N0} rows", dataFont, XBrushes.Gray, leftMargin, yPosition);
                    }
                }
                
                // Save to memory stream and return
                using var stream = new MemoryStream();
                document.Save(stream);
                document.Close();
                gfx.Dispose();
                
                return stream.ToArray();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "PDF generation failed");
                // Fallback to text-based report
                var fallback = $"PDF Generation Error: {ex.Message}\n\nPlease use CSV or Excel export instead.";
                return Encoding.UTF8.GetBytes(fallback);
            }
        }
    }
}