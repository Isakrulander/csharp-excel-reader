# Custom DataFrame Implementation - Dependencies

## NuGet Packages

### EPPlus 8.2.0
- **License**: Polyform Noncommercial (Free for non-commercial use)
- **Purpose**: Excel file processing (.xlsx, .xlsm)
- **Config**: Non-commercial license set in app.config

### iTextSharp.LGPLv2.Core 3.7.7
- **License**: LGPL v2.1 (Open Source)
- **Purpose**: PDF document generation
- **Dependencies**: BouncyCastle.Cryptography, SkiaSharp

### System.Configuration.ConfigurationManager 9.0.9
- **License**: MIT (Microsoft)
- **Purpose**: Configuration file management

## Built-in .NET 8.0 Libraries
- System.IO, System.Text, System.Linq
- System.Collections.Generic, System.Globalization, System.Math

## Installation
```bash
cd CustomDataFrame
dotnet restore
dotnet build
dotnet run
```

## License Notes
- **EPPlus**: Commercial license required for commercial use
- **iTextSharp**: LGPL allows commercial use with attribution
- **Microsoft**: MIT allows commercial use

## Features Provided
- ✅ Excel reading with EPPlus
- ✅ CSV export with custom implementation
- ✅ PDF export with iTextSharp
- ✅ Complete DataFrame operations (filter, sort, stats)
- ✅ Multi-worksheet support