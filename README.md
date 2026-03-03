# PdfBoxPOC – PDF Template Designer & Batch CSV Export

Windows WPF desktop tool for mapping structured PDF data into CSV.

## Features

- Open text-selectable PDFs
- Draw rectangular regions to define Fields and Tables
- Save/Load JSON templates
- Batch process multiple PDFs into a single output.csv
- Undo / Redo
- Visual overlay editing

## Build (Requires .NET 8 SDK)

```
dotnet restore
dotnet publish PdfBoxPOC.csproj -c Release -r win-x64 --self-contained true
```

Published output:
```
bin/Release/net8.0-windows/win-x64/publish/
```

## Usage

1. Open PDF
2. Draw region
3. Add Field / Table / Column
4. Preview Export
5. Save Template
6. Batch → output.csv

## License

MIT
