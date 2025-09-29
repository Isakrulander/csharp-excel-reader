# C# Excel Reader with DataFrame

A C# Excel reader that provides pandas-like DataFrame functionality using EPPlus for Excel file processing.

## Features

- **DataFrame Interface**: Pandas-inspired data structure with rows and columns
- **Multiple Worksheets**: Read specific worksheets by index or name
- **Data Analysis**: Filter, sort, and calculate statistics
- **Excel Export**: Save DataFrame back to Excel files
- **Error Handling**: Comprehensive validation and error messages

## Requirements

- .NET 8.0
- EPPlus 8.2.0 (included)

## Quick Start

```bash
git clone https://github.com/Isakrulander/csharp-excel-reader.git
cd csharp-excel-reader
dotnet run
```

## Usage

```bash
# Use default file (test.xlsx)
dotnet run

# Use specific file
dotnet run myfile.xlsx
```

## Example Output

```
Advanced Excel Reader with DataFrame
Reading Excel file: test.xlsx
============================================================

Worksheets found: 1
  0: 'Sheet1' (4 rows × 2 columns)

Excel data as DataFrame:
============================================================
a       b
------
2       5
3       6
4       7

Shape: 3 rows × 2 columns

Statistics:
  a: Count=3, Mean=3,00, Min=2, Max=4
  b: Count=3, Mean=6,00, Min=5, Max=7

Sorted by 'a' (ascending):
a       b
------
2       5
3       6
4       7

Filtered ('a' > 3,0):
a       b
------
4       7

Exported enhanced DataFrame to: test.enhanced.xlsx
```

## API

### DataFrame
```csharp
// Properties
List<string> Headers
List<Dictionary<string, object>> Rows
int RowCount
int ColumnCount

// Methods
object GetValue(int rowIndex, string columnName)
List<object> GetColumn(string columnName)
void AddRow(Dictionary<string, object> row)
void Display()

// Advanced Operations
DataFrame Filter(Func<Dictionary<string, object>, bool> predicate)
DataFrame SortBy(string columnName, bool ascending = true)
Dictionary<string, double> GetStats(string columnName) // Count, Sum, Mean, Min, Max, StdDev
void ToExcel(string filePath, string worksheetName = "Data")
```

### ExcelReader
```csharp
DataFrame ReadDataFrame(string filePath)
DataFrame ReadDataFrame(string filePath, int worksheetIndex)
DataFrame ReadDataFrame(string filePath, string worksheetName)
List<(string Name, int Index, int Rows, int Columns)> GetWorksheetInfo(string filePath)
```

## Code Examples

```csharp
var reader = new ExcelReader();
var df = reader.ReadDataFrame("data.xlsx");

// Basic operations
df.Display();
var value = df.GetValue(0, "ColumnName");
var column = df.GetColumn("ColumnName");

// Statistics
var stats = df.GetStats("NumericColumn");
// Returns: Count, Sum, Mean, Min, Max, StdDev

// Filtering
var filtered = df.Filter(row => (double)row["Value"] > 10);

// Sorting
var sorted = df.SortBy("ColumnName", ascending: false);

// Export
df.ToExcel("output.xlsx", "MyData");

// Multiple worksheets
var worksheets = reader.GetWorksheetInfo("file.xlsx");
var sheet1 = reader.ReadDataFrame("file.xlsx", 0);
var namedSheet = reader.ReadDataFrame("file.xlsx", "SheetName");
```

## Project Structure

```
├── Program.cs          # DataFrame and ExcelReader classes
├── NKTCS.csproj       # Project configuration
├── README.md          # Documentation
└── test.xlsx          # Sample Excel file
```

## Implementation Details

- **EPPlus 8.2.0** for Excel file processing
- **Non-commercial license** configuration included
- **Single file architecture** for simplicity
- **Error handling** with descriptive messages
- **Data type detection** for dates and numbers

## License

Uses EPPlus under Polyform Noncommercial License for non-commercial use.