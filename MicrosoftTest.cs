public class MicrosoftDataAnalysisTest
{
    public static void TestMicrosoftDataAnalysis(string[] args)
    {
        try
        {
            Console.WriteLine("\n" + new string('=', 60));
            Console.WriteLine("COMPARISON: Microsoft.Data.Analysis DataFrame");
            Console.WriteLine("Reading the SAME test.xlsx file");
            Console.WriteLine(new string('=', 60));

            string excelFilePath = args.Length > 0 ? args[0] : "test.xlsx";
            Console.WriteLine($"Reading Excel file: {excelFilePath}");

            // Read the same Excel file using our ExcelReader, then convert to Microsoft DataFrame
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

            Console.WriteLine("\nMicrosoft.Data.Analysis test with SAME data completed successfully!");

        }
        catch (Exception ex)
        {
            Console.WriteLine($"Microsoft.Data.Analysis test failed: {ex.Message}");
            Console.WriteLine($"Stack trace: {ex.StackTrace}");
        }
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