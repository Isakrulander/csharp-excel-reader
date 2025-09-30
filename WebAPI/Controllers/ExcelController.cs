using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;

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
                
                var result = new
                {
                    fileName = file.FileName,
                    worksheetName = worksheet.Name,
                    rowCount = worksheet.Dimension?.Rows ?? 0,
                    columnCount = worksheet.Dimension?.Columns ?? 0,
                    data = ReadExcelData(worksheet)
                };

                _logger.LogInformation($"Successfully processed Excel file: {file.FileName}");
                return Ok(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error processing Excel file: {file.FileName}");
                return StatusCode(500, new { error = "Error processing Excel file", details = ex.Message });
            }
        }

        private List<Dictionary<string, object>> ReadExcelData(ExcelWorksheet worksheet)
        {
            var data = new List<Dictionary<string, object>>();
            
            if (worksheet.Dimension == null)
                return data;

            // Read headers from first row
            var headers = new List<string>();
            for (int col = 1; col <= worksheet.Dimension.Columns; col++)
            {
                headers.Add(worksheet.Cells[1, col].Text ?? $"Column{col}");
            }

            // Read data rows (skip header)
            for (int row = 2; row <= worksheet.Dimension.Rows; row++)
            {
                var rowData = new Dictionary<string, object>();
                for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                {
                    var cellValue = worksheet.Cells[row, col].Text ?? "";
                    rowData[headers[col - 1]] = cellValue;
                }
                data.Add(rowData);
            }

            return data;
        }
    }
}