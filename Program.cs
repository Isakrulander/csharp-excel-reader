using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

// Simple DataFrame-like class for C#
public class DataFrame
{
    public List<string> Headers { get; set; }
    public List<Dictionary<string, object>> Rows { get; set; }

    public DataFrame()
    {
        Headers = new List<string>();
        Rows = new List<Dictionary<string, object>>();
    }

    public void AddRow(Dictionary<string, object> row)
    {
        Rows.Add(row);
    }

    public object? GetValue(int rowIndex, string columnName)
    {
        if (rowIndex < Rows.Count && Rows[rowIndex].ContainsKey(columnName))
            return Rows[rowIndex][columnName];
        return null;
    }

    public List<object?> GetColumn(string columnName)
    {
        return Rows.Select(row => row.ContainsKey(columnName) ? row[columnName] : null).ToList();
    }

    /// <summary>
    /// Gets the number of rows in the DataFrame
    /// </summary>
    public int RowCount => Rows.Count;

    /// <summary>
    /// Gets the number of columns in the DataFrame
    /// </summary>
    public int ColumnCount => Headers.Count;

    /// <summary>
    /// Filters the DataFrame based on a predicate
    /// </summary>
    public DataFrame Filter(Func<Dictionary<string, object>, bool> predicate)
    {
        var filtered = new DataFrame();
        filtered.Headers.AddRange(Headers);
        
        foreach (var row in Rows.Where(predicate))
        {
            filtered.AddRow(new Dictionary<string, object>(row));
        }
        
        return filtered;
    }

    /// <summary>
    /// Sorts the DataFrame by a column
    /// </summary>
    public DataFrame SortBy(string columnName, bool ascending = true)
    {
        var sorted = new DataFrame();
        sorted.Headers.AddRange(Headers);

        var sortedRows = ascending 
            ? Rows.OrderBy(row => row.ContainsKey(columnName) ? row[columnName] : null)
            : Rows.OrderByDescending(row => row.ContainsKey(columnName) ? row[columnName] : null);

        foreach (var row in sortedRows)
        {
            sorted.AddRow(new Dictionary<string, object>(row));
        }

        return sorted;
    }

    /// <summary>
    /// Gets basic statistics for a numeric column
    /// </summary>
    public Dictionary<string, double> GetStats(string columnName)
    {
        var values = GetColumn(columnName)
            .Where(v => v != null && IsNumeric(v))
            .Select(v => Convert.ToDouble(v))
            .ToList();

        if (!values.Any())
            throw new InvalidOperationException($"No numeric values found in column '{columnName}'");

        var mean = values.Average();
        var variance = values.Sum(v => Math.Pow(v - mean, 2)) / values.Count;

        return new Dictionary<string, double>
        {
            ["Count"] = values.Count,
            ["Sum"] = values.Sum(),
            ["Mean"] = mean,
            ["Min"] = values.Min(),
            ["Max"] = values.Max(),
            ["StdDev"] = Math.Sqrt(variance)
        };
    }

    /// <summary>
    /// Exports DataFrame to Excel file
    /// </summary>
    public void ToExcel(string filePath, string worksheetName = "Data")
    {
        using var package = new ExcelPackage();
        var worksheet = package.Workbook.Worksheets.Add(worksheetName);

        // Write headers
        for (int col = 0; col < Headers.Count; col++)
        {
            worksheet.Cells[1, col + 1].Value = Headers[col];
            worksheet.Cells[1, col + 1].Style.Font.Bold = true;
        }

        // Write data
        for (int row = 0; row < Rows.Count; row++)
        {
            for (int col = 0; col < Headers.Count; col++)
            {
                var value = Rows[row].ContainsKey(Headers[col]) ? Rows[row][Headers[col]] : null;
                worksheet.Cells[row + 2, col + 1].Value = value;
            }
        }

        // Auto-fit columns
        worksheet.Cells.AutoFitColumns();

        // Save file
        var file = new FileInfo(filePath);
        package.SaveAs(file);
    }

    /// <summary>
    /// Helper method to check if a value is numeric
    /// </summary>
    private static bool IsNumeric(object value)
    {
        return value is sbyte or byte or short or ushort or int or uint or long or ulong or float or double or decimal;
    }

    public void Display()
    {
        if (Headers.Count == 0)
        {
            Console.WriteLine("DataFrame is empty.");
            return;
        }

        // Display headers
        Console.WriteLine(string.Join("\t", Headers));
        Console.WriteLine(new string('-', Headers.Sum(h => h.Length) + (Headers.Count - 1) * 4));

        // Display rows
        foreach (var row in Rows)
        {
            var values = Headers.Select(header => 
                row.ContainsKey(header) ? (row[header]?.ToString() ?? "null") : "null");
            Console.WriteLine(string.Join("\t", values));
        }

        Console.WriteLine($"\nShape: {Rows.Count} rows × {Headers.Count} columns");
    }
}

public class ExcelReader
{
    static ExcelReader()
    {
        // Set EPPlus license for non-commercial personal use (EPPlus 8+)
        ExcelPackage.License.SetNonCommercialPersonal("Personal Use");
    }

    /// <summary>
    /// Gets information about all worksheets in an Excel file
    /// </summary>
    public List<(string Name, int Index, int Rows, int Columns)> GetWorksheetInfo(string filePath)
    {
        if (string.IsNullOrEmpty(filePath))
            throw new ArgumentException("File path cannot be null or empty", nameof(filePath));
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"Excel file not found: {filePath}");
        
        var worksheets = new List<(string Name, int Index, int Rows, int Columns)>();
        
        using var package = new ExcelPackage(new FileInfo(filePath));
        for (int i = 0; i < package.Workbook.Worksheets.Count; i++)
        {
            var ws = package.Workbook.Worksheets[i];
            var rows = ws.Dimension?.Rows ?? 0;
            var cols = ws.Dimension?.Columns ?? 0;
            worksheets.Add((ws.Name, i, rows, cols));
        }
        
        return worksheets;
    }

    /// <summary>
    /// Reads a specific worksheet by index
    /// </summary>
    public DataFrame ReadDataFrame(string filePath, int worksheetIndex)
    {
        return ReadDataFrameInternal(filePath, worksheetIndex);
    }

    /// <summary>
    /// Reads a specific worksheet by name
    /// </summary>
    public DataFrame ReadDataFrame(string filePath, string worksheetName)
    {
        return ReadDataFrameInternal(filePath, worksheetName: worksheetName);
    }

    public DataFrame ReadDataFrame(string filePath)
    {
        return ReadDataFrameInternal(filePath);
    }

    private DataFrame ReadDataFrameInternal(string filePath, int? worksheetIndex = null, string? worksheetName = null)
    {
        if (string.IsNullOrEmpty(filePath))
            throw new ArgumentException("File path cannot be null or empty", nameof(filePath));
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"Excel file not found: {filePath}");

        var dataFrame = new DataFrame();

        try
        {
            using var package = new ExcelPackage(new FileInfo(filePath));
            
            if (package.Workbook.Worksheets.Count == 0)
            {
                Console.WriteLine("The Excel file contains no worksheets.");
                return dataFrame;
            }

            // Select worksheet
            ExcelWorksheet worksheet;
            if (!string.IsNullOrEmpty(worksheetName))
            {
                worksheet = package.Workbook.Worksheets[worksheetName];
                if (worksheet == null)
                    throw new ArgumentException($"Worksheet '{worksheetName}' not found");
            }
            else
            {
                var index = worksheetIndex ?? 0;
                if (index >= package.Workbook.Worksheets.Count)
                    throw new ArgumentException($"Worksheet index {index} is out of range");
                worksheet = package.Workbook.Worksheets[index];
            }

            if (worksheet.Dimension == null)
            {
                Console.WriteLine($"The worksheet '{worksheet.Name}' is empty.");
                return dataFrame;
            }

            Console.WriteLine($"Reading worksheet '{worksheet.Name}' ({worksheet.Dimension.Rows} rows, {worksheet.Dimension.Columns} columns)");

            int rowCount = worksheet.Dimension.Rows;
            int colCount = worksheet.Dimension.Columns;

            // Read headers from first row with better handling
            for (int col = 1; col <= colCount; col++)
            {
                var headerValue = worksheet.Cells[1, col].Value?.ToString()?.Trim();
                if (string.IsNullOrEmpty(headerValue))
                    headerValue = $"Column{col}";
                    
                // Ensure unique headers
                var uniqueHeader = EnsureUniqueHeader(dataFrame.Headers, headerValue);
                dataFrame.Headers.Add(uniqueHeader);
            }

            // Read data rows (starting from row 2) with better data type handling
            for (int row = 2; row <= rowCount; row++)
            {
                var rowData = new Dictionary<string, object>();
                bool hasData = false;

                for (int col = 1; col <= colCount; col++)
                {
                    var cell = worksheet.Cells[row, col];
                    var cellValue = ProcessCellValue(cell);
                    var headerName = dataFrame.Headers[col - 1];
                    
                    rowData[headerName] = cellValue ?? "";
                    if (cellValue != null) hasData = true;
                }

                // Only add rows with data
                if (hasData)
                {
                    dataFrame.AddRow(rowData);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while reading Excel: {ex.Message}");
            throw;
        }

        return dataFrame;
    }

    /// <summary>
    /// Processes cell value with proper data type handling
    /// </summary>
    private static object? ProcessCellValue(ExcelRange cell)
    {
        var value = cell.Value;
        if (value == null) return null;

        // Handle dates
        if (value is double doubleValue && IsDateFormat(cell.Style.Numberformat.Format))
        {
            try
            {
                return DateTime.FromOADate(doubleValue);
            }
            catch
            {
                return doubleValue;
            }
        }

        return value;
    }

    /// <summary>
    /// Checks if number format represents a date
    /// </summary>
    private static bool IsDateFormat(string format)
    {
        return !string.IsNullOrEmpty(format) && 
               (format.Contains("d") || format.Contains("m") || format.Contains("y") || 
                format.Contains("h") || format.Contains("s"));
    }

    /// <summary>
    /// Ensures header names are unique
    /// </summary>
    private static string EnsureUniqueHeader(List<string> existingHeaders, string proposedName)
    {
        if (!existingHeaders.Contains(proposedName))
            return proposedName;

        int counter = 1;
        string uniqueName;
        do
        {
            uniqueName = $"{proposedName}_{counter}";
            counter++;
        } while (existingHeaders.Contains(uniqueName));

        return uniqueName;
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        try
        {
            // Use command line argument or default file
            string excelFilePath = args.Length > 0 ? args[0] : "test.xlsx";
            
            Console.WriteLine("Advanced Excel Reader with DataFrame");
            Console.WriteLine($"Reading Excel file: {excelFilePath}");
            Console.WriteLine(new string('=', 60));

            // Create an instance of the ExcelReader class
            var reader = new ExcelReader();

            // Show worksheet information  
            var worksheetInfo = reader.GetWorksheetInfo(excelFilePath);
            Console.WriteLine($"\nWorksheets found: {worksheetInfo.Count}");
            foreach (var (name, index, rows, cols) in worksheetInfo)
            {
                Console.WriteLine($"  {index}: '{name}' ({rows} rows × {cols} columns)");
            }

            // Read the data as a DataFrame
            var dataFrame = reader.ReadDataFrame(excelFilePath);

            // Display DataFrame (pandas-like format)
            Console.WriteLine("\nExcel data as DataFrame:");
            Console.WriteLine(new string('=', 60));
            dataFrame.Display();

            // Show basic DataFrame info
            Console.WriteLine("\nDataFrame Info:");
            Console.WriteLine($"- Shape: {dataFrame.RowCount} rows × {dataFrame.ColumnCount} columns");
            Console.WriteLine($"- Columns: [{string.Join(", ", dataFrame.Headers)}]");

            // Demonstrate advanced features
            if (dataFrame.Headers.Count > 0 && dataFrame.RowCount > 0)
            {
                Console.WriteLine("\nAdvanced DataFrame Features:");
                
                // Statistics for numeric columns
                Console.WriteLine("\nStatistics:");
                foreach (var column in dataFrame.Headers)
                {
                    try
                    {
                        var stats = dataFrame.GetStats(column);
                        Console.WriteLine($"  {column}: Count={stats["Count"]}, Mean={stats["Mean"]:F2}, Min={stats["Min"]}, Max={stats["Max"]}");
                    }
                    catch (InvalidOperationException)
                    {
                        Console.WriteLine($"  {column}: Non-numeric column");
                    }
                }

                // Sorting example
                var firstColumn = dataFrame.Headers[0];
                Console.WriteLine($"\nSorted by '{firstColumn}' (ascending):");
                var sorted = dataFrame.SortBy(firstColumn);
                sorted.Display();

                // Filtering example (show only rows where first column > median)
                try
                {
                    var stats = dataFrame.GetStats(firstColumn);
                    var median = stats["Mean"]; // Using mean as approximate median
                    var filtered = dataFrame.Filter(row => 
                        row.ContainsKey(firstColumn) && 
                        row[firstColumn] is double value && 
                        value > median);
                    
                    if (filtered.RowCount > 0)
                    {
                        Console.WriteLine($"\nFiltered ('{firstColumn}' > {median:F1}):");
                        filtered.Display();
                    }
                }
                catch (Exception)
                {
                    Console.WriteLine($"\nFiltering: Cannot filter non-numeric column '{firstColumn}'");
                }

                // Export example
                var exportPath = Path.ChangeExtension(excelFilePath, ".enhanced.xlsx");
                dataFrame.ToExcel(exportPath, "Enhanced Data");
                Console.WriteLine($"\nExported enhanced DataFrame to: {exportPath}");

                Console.WriteLine("\nAll advanced features demonstrated successfully!");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
            Console.WriteLine("\nUsage: dotnet run [excel-file-path]");
            Console.WriteLine("Example: dotnet run myfile.xlsx");
            Environment.Exit(1);
        }
    }
}