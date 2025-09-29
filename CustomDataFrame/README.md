# Custom DataFrame Implementation

Complete custom DataFrame implementation with full Excel workflow support.

## 🚀 Quick Start

```bash
dotnet run
```

## ✨ Features

- ✅ **Complete Excel Support**: Read Excel files with EPPlus
- ✅ **Multi-Format Export**: Excel, CSV, and PDF output
- ✅ **DataFrame Operations**: Filter, sort, statistics
- ✅ **Multi-Worksheet**: Read specific sheets by name/index
- ✅ **Custom Implementation**: Full control over functionality

## 📊 Sample Output

```
Custom DataFrame Implementation
============================================================
Excel data as DataFrame:
a       b
------
2       5
3       6
4       7

Statistics:
  a: Count=3, Mean=3.00, Min=2, Max=4
  b: Count=3, Mean=6.00, Min=5, Max=7

✅ Exported to Excel, CSV, and PDF
```

## 🛠 Dependencies

- **EPPlus 8.2.0** - Excel processing
- **iTextSharp.LGPLv2.Core 3.7.7** - PDF generation
- **System.Configuration.ConfigurationManager 9.0.9** - Configuration

## 📝 Usage

```bash
# Use default test.xlsx
dotnet run

# Use your Excel file
dotnet run path/to/your/file.xlsx
```

## 🎯 Best For

- Excel processing workflows
- Complete export capabilities needed
- Educational purposes
- Small to medium datasets (<10,000 rows)
- When you want full control over implementation

See `DEPENDENCIES.md` and `requirements.txt` for detailed dependency information.