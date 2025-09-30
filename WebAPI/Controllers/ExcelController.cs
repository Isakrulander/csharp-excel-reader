using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using Microsoft.Data.Analysis;
using System.Globalization;

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
    }
}