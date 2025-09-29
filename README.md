# C# Excel Reader with DataFrame

A C# Excel reader that provides pandas-like DataFrame functionality using EPPlus for Excel file processing.

## Features

- **DataFrame Interface**: Pandas-inspired data structure with rows and columns
- **Multiple Worksheets**: Read specific worksheets by index or name
- **Data Analysis**: Filter, sort, and calculate statistics
- **Multi-Format Export**: Save DataFrame to Excel (.xlsx), CSV (.csv), and PDF (.pdf)
- **Advanced Statistics**: Count, mean, min, max, standard deviation
- **Data Type Handling**: Automatic detection of numbers, dates, and text
- **Error Handling**: Comprehensive validation and error messages

## Requirements

### System Requirements
- .NET 8.0 SDK or later
- Windows, macOS, or Linux

### NuGet Dependencies
- **EPPlus 8.2.0** - Excel file processing (.xlsx, .xlsm)
- **iTextSharp.LGPLv2.Core 3.7.7** - PDF document generation
- **System.Configuration.ConfigurationManager 9.0.9** - Configuration management

### Installation
```bash
# Clone the repository
git clone https://github.com/Isakrulander/csharp-excel-reader.git
cd csharp-excel-reader

# Restore NuGet packages
dotnet restore

# Build the project
dotnet build
```

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

Exported enhanced DataFrame to Excel: test.enhanced.xlsx
Exported enhanced DataFrame to CSV: test.enhanced.csv
Exported enhanced DataFrame to PDF: test.enhanced.pdf
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
void ToCsv(string filePath, string delimiter = ",")
void ToPdf(string filePath, string title = "DataFrame Report")
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

// Export to multiple formats
df.ToExcel("output.xlsx", "MyData");
df.ToCsv("output.csv");
df.ToPdf("report.pdf", "Sales Report");

// Multiple worksheets
var worksheets = reader.GetWorksheetInfo("file.xlsx");
var sheet1 = reader.ReadDataFrame("file.xlsx", 0);
var namedSheet = reader.ReadDataFrame("file.xlsx", "SheetName");
```

## Project Structure

```
├── Program.cs          # DataFrame and ExcelReader classes
├── NKTCS.csproj       # Project configuration and NuGet packages
├── app.config         # EPPlus license configuration
├── README.md          # Project documentation
├── requirements.txt   # Dependency list with descriptions
├── DEPENDENCIES.md    # Detailed library documentation
└── test.xlsx          # Sample Excel file
```

## Libraries Used

### EPPlus 8.2.0
- **Purpose**: Excel file reading and writing (.xlsx, .xlsm formats)
- **License**: Polyform Noncommercial 1.0.0 (Free for non-commercial use)
- **Features Used**: Workbook manipulation, worksheet access, cell reading/writing, auto-fit columns
- **Configuration**: Non-commercial license set in app.config

### iTextSharp.LGPLv2.Core 3.7.7
- **Purpose**: PDF document generation and manipulation
- **License**: LGPL v2 (Open Source)
- **Features Used**: PDF table creation, text formatting, document layout, cell styling
- **Dependencies**: BouncyCastle.Cryptography (2.6.2), SkiaSharp (3.119.0)

### System.Configuration.ConfigurationManager 9.0.9
- **Purpose**: Application configuration management
- **License**: MIT (Microsoft)
- **Features Used**: Reading app.config file for EPPlus license configuration

### Built-in .NET Libraries
- **System.IO**: File operations and stream handling
- **System.Text**: Text encoding (UTF-8) and StringBuilder for CSV generation
- **System.Linq**: LINQ operations for data filtering and transformation
- **System.Collections.Generic**: Generic collections (List, Dictionary)
- **System.Globalization**: Number and date formatting

## Implementation Details

- **Single file architecture** for simplicity (Program.cs)
- **Pandas-inspired API** design for familiar data manipulation
- **Comprehensive error handling** with descriptive messages
- **Automatic data type detection** for numbers, dates, and text
- **Memory efficient** streaming for large datasets
- **Cross-platform compatible** (.NET 8.0)

## License

- **EPPlus**: Polyform Noncommercial License (Free for non-commercial use)
- **iTextSharp**: LGPL v2 (Open Source)
- **Microsoft Libraries**: MIT License
- **Project Code**: Open Source (specify your preferred license)