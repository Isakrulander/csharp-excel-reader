# Microsoft.Data.Analysis Dependencies

Professional DataFrame implementation for corporate environments.

## Core NuGet Packages

### Microsoft.Data.Analysis 0.21.1
- **License**: MIT (Microsoft Open Source)
- **Purpose**: Professional DataFrame operations with pandas-like API
- **Features**: Vectorized operations, advanced statistics, high performance
- **Maintained by**: Microsoft Corporation

### EPPlus 8.2.0
- **License**: Polyform Noncommercial (Free for non-commercial use)
- **Purpose**: Excel file processing (.xlsx, .xlsm)
- **Configuration**: Non-commercial license set in app.config
- **Commercial Use**: Requires commercial license

### iTextSharp.LGPLv2.Core 3.7.7
- **License**: LGPL v2.1 (Open Source)
- **Purpose**: Professional PDF document generation
- **Dependencies**: BouncyCastle.Cryptography, SkiaSharp
- **Commercial Use**: LGPL allows commercial use with attribution

### System.Configuration.ConfigurationManager 9.0.9
- **License**: MIT (Microsoft)
- **Purpose**: Application configuration management
- **Features**: app.config reading, EPPlus license configuration

## Built-in .NET 8.0 Libraries
- System.IO, System.Text, System.Linq
- System.Collections.Generic, System.Globalization
- Microsoft.Data.Analysis extensions

## Installation
```bash
cd MicrosoftDataAnalysis
dotnet restore
dotnet build
dotnet run
```

## Corporate License Compliance

### For Commercial Use:
- **Microsoft.Data.Analysis**: ✅ MIT license allows commercial use
- **EPPlus**: ⚠️ Requires commercial license for commercial use
- **iTextSharp**: ✅ LGPL allows commercial use with proper attribution
- **System.Configuration.ConfigurationManager**: ✅ MIT allows commercial use

### For Non-Commercial Use:
- All libraries are free to use for educational, personal, and open-source projects

## Performance Characteristics
- **Optimized for**: Large datasets (>50,000 rows)
- **Memory**: Efficient columnar storage
- **Processing**: Vectorized operations using SIMD
- **Scalability**: Enterprise-grade performance