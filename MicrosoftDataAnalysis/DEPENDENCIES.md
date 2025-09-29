# Microsoft.Data.Analysis Implementation - Dependencies

## NuGet Packages

### EPPlus 8.2.0
- **License**: Polyform Noncommercial (Free for non-commercial use)
- **Purpose**: Excel file processing for data conversion
- **Config**: Non-commercial license set in app.config

### Microsoft.Data.Analysis 0.21.1
- **License**: MIT (Microsoft Open Source)
- **Purpose**: Professional DataFrame operations
- **Features**: Pandas-like API, vectorized operations, high performance

### iTextSharp.LGPLv2.Core 3.7.7
- **License**: LGPL v2.1 (Open Source)
- **Purpose**: PDF document generation
- **Dependencies**: BouncyCastle.Cryptography, SkiaSharp

### System.Configuration.ConfigurationManager 9.0.9
- **License**: MIT (Microsoft)
- **Purpose**: Configuration file management

## Built-in .NET 8.0 Libraries
- System.IO, System.Linq, System.Collections.Generic, System.Text
- Microsoft.Data.Analysis extensions

## Installation
```bash
cd MicrosoftDataAnalysis
dotnet restore
dotnet build
dotnet run
```

## License Notes
- **EPPlus**: Commercial license required for commercial use
- **Microsoft.Data.Analysis**: MIT allows commercial use
- **Microsoft Libraries**: MIT allows commercial use

## Features Provided
- ✅ Excel reading via EPPlus conversion
- ✅ Excel export with professional formatting
- ✅ Advanced statistical operations
- ✅ Pandas-like vectorized filtering
- ✅ High-performance columnar storage
- ✅ Professional Microsoft-backed library
- ✅ CSV export functionality
- ✅ PDF export with professional formatting