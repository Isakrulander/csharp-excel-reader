# C# Excel Reader with DataFrame

A simple, pandas-like Excel reader for C# that uses EPPlus to read Excel files and display them in a structured DataFrame format.

## Features

- 🐼 **Pandas-like DataFrame**: Familiar interface for data manipulation
- 📊 **Excel Reading**: Supports .xlsx files with automatic header detection
- 🔍 **Data Access**: Get values by row/column, extract entire columns
- 📝 **Pretty Display**: Formatted table output in the console
- ⚡ **Fast**: Built with EPPlus for efficient Excel processing
- 🛡️ **Error Handling**: Comprehensive error messages and validation

## Requirements

- .NET 8.0 or later
- EPPlus 8.2.0 (automatically installed)

## Installation

1. Clone the repository:
```bash
git clone https://github.com/Isakrulander/csharp-excel-reader.git
cd csharp-excel-reader
```

2. Restore dependencies:
```bash
dotnet restore
```

## Usage

### Basic Usage

Run with the default Excel file (`test.xlsx`):
```bash
dotnet run
```

### Specify Excel File

Run with a specific Excel file:
```bash
dotnet run path/to/your/file.xlsx
```

### Example Output

```
Reading Excel file: test.xlsx
Excel data as DataFrame:
=========================================
a       b
------
2       5
3       6
4       7

Shape: 3 rows × 2 columns

--- DataFrame functions ---
Column names: [a, b]

Values in column 'a': [2, 3, 4]
First value in 'a': 2
```

## Code Examples

### Basic DataFrame Operations

```csharp
var reader = new ExcelReader();
var dataFrame = reader.ReadDataFrame("data.xlsx");

// Display entire DataFrame
dataFrame.Display();

// Get column names
Console.WriteLine(string.Join(", ", dataFrame.Headers));

// Get specific column values
var columnValues = dataFrame.GetColumn("ColumnName");

// Get specific cell value
var cellValue = dataFrame.GetValue(0, "ColumnName"); // First row
```

### Error Handling

The application includes comprehensive error handling:
- File not found errors
- Invalid file format errors
- Null/empty parameter validation
- Informative error messages

## Project Structure

```
├── Program.cs          # Main application and DataFrame implementation
├── NKTCS.csproj       # Project configuration and dependencies
├── .gitignore         # Git ignore rules
├── app.config         # EPPlus configuration
└── README.md          # This file
```

## Dependencies

- **EPPlus** (8.2.0): Excel file processing
- **System.Configuration.ConfigurationManager** (9.0.9): Configuration management

## License

This project uses EPPlus under the **Polyform Noncommercial License** for non-commercial use. 

For commercial use, you need to purchase a commercial EPPlus license from [EPPlus Software](https://epplussoftware.com/).

## Development

### Building the Project

```bash
dotnet build
```

### Running Tests

```bash
dotnet test  # (when tests are added)
```

## Features Roadmap

- [ ] **Multiple Worksheets**: Support for reading specific worksheets
- [ ] **Data Types**: Better handling of dates, numbers, and formulas
- [ ] **Export Features**: Save DataFrame back to Excel
- [ ] **Filtering**: Add data filtering capabilities
- [ ] **Sorting**: Add data sorting functionality
- [ ] **Statistics**: Basic statistical operations (sum, mean, count)
- [ ] **Unit Tests**: Comprehensive test coverage

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/new-feature`)
3. Commit your changes (`git commit -am 'Add new feature'`)
4. Push to the branch (`git push origin feature/new-feature`)
5. Create a Pull Request

## Error Handling

The application provides clear error messages for common issues:

- **File not found**: `Excel file not found: filename.xlsx`
- **Invalid arguments**: Usage instructions with examples
- **Excel format errors**: Detailed error descriptions

## Performance

- Efficiently handles large Excel files
- Memory-conscious DataFrame implementation
- Fast Excel parsing with EPPlus

## Examples

### Sample Excel File Structure

Your Excel file should have headers in the first row:

| Name    | Age | City      |
|---------|-----|-----------|
| Alice   | 25  | Stockholm |
| Bob     | 30  | Göteborg  |
| Charlie | 35  | Malmö     |

### Expected Output

```
Excel data as DataFrame:
=========================================
Name    Age     City
-----------------
Alice   25      Stockholm
Bob     30      Göteborg
Charlie 35      Malmö

Shape: 3 rows × 3 columns
```

---

**Made with ❤️ using C# and EPPlus**