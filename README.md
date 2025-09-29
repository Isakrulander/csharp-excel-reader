# C# Excel Reader - DataFrame Implementations

Two different approaches to implementing pandas-like DataFrame functionality in C#.

## 🚀 Quick Start

Choose your implementation:

### Custom DataFrame Implementation (Full-Featured)
```bash
cd CustomDataFrame
dotnet run
```

### Microsoft.Data.Analysis Implementation  
```bash
cd MicrosoftDataAnalysis
dotnet run
```

## 📁 Project Structure

```
├── CustomDataFrame/           # Complete custom implementation
│   ├── Program.cs            # Custom DataFrame with Excel/CSV/PDF export
│   ├── CustomDataFrame.csproj
│   ├── README.md             # Folder-specific documentation
│   ├── DEPENDENCIES.md       # Detailed dependencies
│   ├── requirements.txt      # Package requirements
│   ├── test.xlsx             # Local test file
│   └── app.config            # EPPlus configuration
├── MicrosoftDataAnalysis/     # Microsoft's official library
│   ├── Program.cs            # Uses Microsoft.Data.Analysis
│   ├── MicrosoftDataAnalysis.csproj
│   ├── README.md             # Folder-specific documentation
│   ├── DEPENDENCIES.md       # Detailed dependencies
│   ├── requirements.txt      # Package requirements
│   ├── test.xlsx             # Local test file
│   └── app.config            # EPPlus configuration
├── README.md                  # This file (main overview)
├── COMPARISON.md             # Detailed comparison
├── DEPENDENCIES.md           # Overall dependencies info
└── requirements.txt          # General dependencies overview
```

## 🎯 Key Differences

| Feature | Custom DataFrame | Microsoft.Data.Analysis |
|---------|------------------|-------------------------|
| **Excel Reading** | ✅ Built-in | ✅ Via conversion |
| **Excel Export** | ✅ Custom | ✅ Professional |
| **CSV Export** | ✅ Custom | ✅ Corporate-grade |
| **PDF Export** | ✅ With iTextSharp | ✅ Professional reports |
| **Statistics** | ✅ Basic stats | ✅ Advanced operations |
| **Filtering** | ✅ Lambda expressions | ✅ Vectorized |
| **Performance** | Good for <10K rows | Optimized for >50K rows |
| **Best For** | Learning/Simple workflows | Corporate/Teams |

## 📊 Sample Results

Both process the same data identically:

```
Excel data as DataFrame:
a       b
------
2       5
3       6
4       7

Statistics:
  a: Count=3, Mean=3.00, Min=2, Max=4
  b: Count=3, Mean=6.00, Min=5, Max=7
```

**Custom Implementation**: ✅ Complete Excel workflow + Custom exports  
**Microsoft Implementation**: ✅ Professional DataFrame + Complete export suite (Excel/CSV/PDF)

## 🏆 Recommendations

**Use CustomDataFrame for:**
- Excel processing workflows
- Need multiple export formats
- Want complete control
- Small to medium datasets

**Use MicrosoftDataAnalysis for:**
- **Corporate/Enterprise environments** 🏢
- Large datasets (>50K rows)
- Advanced data science operations
- Team development (standard library)
- Performance-critical applications
- Professional reporting requirements

## 🛠 Dependencies

Both use EPPlus for Excel reading. Custom adds iTextSharp for PDF export.

## 📝 Usage

```bash
# Run with default test.xlsx
dotnet run

# Run with your Excel file
dotnet run path/to/your/file.xlsx
```

Both implementations produce identical analytical results - choose based on your workflow needs!