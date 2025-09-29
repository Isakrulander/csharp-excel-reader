using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

/// <summary>
/// A simple DataFrame-like class for C# that mimics pandas DataFrame functionality
/// </summary>
public class DataFrame
{
    public List<string> Headers { get; set; }
    public List<Dictionary<string, object>> Rows { get; set; }

    public DataFrame()
    {
        Headers = new List<string>();
        Rows = new List<Dictionary<string, object>>();
    }

    /// <summary>
    /// Adds a new row to the DataFrame
    /// </summary>
    /// <param name="row">Dictionary containing column names and values</param>
    public void AddRow(Dictionary<string, object> row)
    {
        if (row == null)
            throw new ArgumentNullException(nameof(row), "Row cannot be null");
        
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
    /// Displays the DataFrame in a formatted table
    /// </summary>
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

/// <summary>
/// Service for reading Excel files and converting them to DataFrame objects
/// </summary>
public class ExcelReader
{
    static ExcelReader()
    {
        // Set EPPlus license for non-commercial personal use (EPPlus 8+)
        ExcelPackage.License.SetNonCommercialPersonal("Personal Use");
    }

    /// <summary>
    /// Reads an Excel file and returns the data as a DataFrame
    /// </summary>
    /// <param name="filePath">Path to the Excel file</param>
    /// <returns>DataFrame containing the Excel data</returns>
    public DataFrame ReadDataFrame(string filePath)
    {
        // Validate input
        if (string.IsNullOrEmpty(filePath))
            throw new ArgumentException("File path cannot be null or empty", nameof(filePath));

        if (!File.Exists(filePath))
            throw new FileNotFoundException($"Excel file not found: {filePath}");

        var dataFrame = new DataFrame();

        try
        {
            var file = new FileInfo(filePath);

            using (var package = new ExcelPackage(file))
            {
                var worksheet = package.Workbook.Worksheets[0];

                if (worksheet.Dimension == null)
                {
                    Console.WriteLine("The worksheet is empty.");
                    return dataFrame;
                }

                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                // Read headers from first row
                for (int col = 1; col <= colCount; col++)
                {
                    var headerValue = worksheet.Cells[1, col].Value?.ToString() ?? $"Column{col}";
                    dataFrame.Headers.Add(headerValue);
                }

                // Read data rows (starting from row 2)
                for (int row = 2; row <= rowCount; row++)
                {
                    var rowData = new Dictionary<string, object>();
                    for (int col = 1; col <= colCount; col++)
                    {
                        var cellValue = worksheet.Cells[row, col].Value;
                        var headerName = dataFrame.Headers[col - 1];
                        rowData[headerName] = cellValue;
                    }
                    dataFrame.AddRow(rowData);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }

        return dataFrame;
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
            
            Console.WriteLine($"Reading Excel file: {excelFilePath}");

        // Create an instance of the ExcelReader class
        var reader = new ExcelReader();

        // Read the data as a DataFrame
        var dataFrame = reader.ReadDataFrame(excelFilePath);

        // Display DataFrame (pandas-like format)
        Console.WriteLine("Excel data as DataFrame:");
        Console.WriteLine("=" + new string('=', 40));
        dataFrame.Display();

        // Example DataFrame functions
        if (dataFrame.Headers.Count > 0)
        {
            Console.WriteLine("\n--- DataFrame functions ---");
            Console.WriteLine($"Column names: [{string.Join(", ", dataFrame.Headers)}]");
            
            // Show first column values
            var firstColumn = dataFrame.Headers[0];
            var columnValues = dataFrame.GetColumn(firstColumn);
            Console.WriteLine($"\nValues in column '{firstColumn}': [{string.Join(", ", columnValues)}]");
            
            // Show specific cell
            if (dataFrame.Rows.Count > 0)
            {
                var firstCellValue = dataFrame.GetValue(0, firstColumn);
                Console.WriteLine($"First value in '{firstColumn}': {firstCellValue}");
            }
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