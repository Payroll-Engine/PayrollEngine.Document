# Payroll Engine Document

> Part of the [Payroll Engine](https://github.com/Payroll-Engine/PayrollEngine) open-source payroll automation framework.
> Full documentation at [payrollengine.org](https://payrollengine.org).

Document library for the Payroll Engine, providing data transformation for exchange and reporting. Supports PDF generation via FastReport templates, Excel export via NPOI, and FastReport template generation from a DataSet schema.

## Features
- **PDF Generation** — Merge `DataSet` with FastReport templates (`.frx`) to produce PDF documents with metadata (author, title, keywords)
- **Excel Export** — Convert `DataSet` tables to `.xlsx` workbooks with auto-sized columns, header filters, and correct cell typing (numeric, boolean, date)
- **Template Generation** — Generate a FastReport skeleton (`.frx`) from a DataSet schema for use in the FastReport Designer, or update the Dictionary section of an existing template in CI pipelines
- **Template Parameters** — Pass custom key-value parameters into FastReport templates
- **Document Metadata** — Apply metadata (author, title, subject, keywords) to generated documents

---

## NuGet Package

Available on [NuGet.org](https://www.nuget.org/profiles/PayrollEngine):

```sh
dotnet add package PayrollEngine.Document
```

---

## Quick Start

### PDF Generation from FastReport Template

```csharp
using PayrollEngine.Document;

IDocumentService documentService = new DocumentService();

// load your FastReport template (.frx)
await using var templateStream = File.OpenRead("report-template.frx");

// prepare your data
var dataSet = new DataSet();
// ... populate tables ...

// define document metadata
var metadata = new DocumentMetadata
{
    Title = "Payroll Report Q1 2026",
    Author = "Payroll Engine",
    Keywords = "payroll, report"
};

// optional: template parameters
var parameters = new Dictionary<string, object>
{
    ["CompanyName"] = "Acme Corp",
    ["ReportDate"] = DateTime.Now
};

// generate PDF
await using var pdfStream = await documentService.MergeAsync(
    templateStream, dataSet, DocumentType.Pdf, metadata, parameters);

// save to file
await pdfStream.WriteToFile("report.pdf");
```

### Excel Export

```csharp
IDocumentService documentService = new DocumentService();

var dataSet = new DataSet();
// ... populate tables (each DataTable becomes a worksheet) ...

var metadata = new DocumentMetadata { Title = "Payroll Data Export" };

// generate Excel workbook
await using var excelStream = await documentService.ExcelMergeAsync(dataSet, metadata);

// save to file
await excelStream.WriteToFile("export.xlsx");
```

### FastReport Template Generation

```csharp
IDocumentService documentService = new DocumentService();

var dataSet = new DataSet();
// ... populate tables and relations ...

// new skeleton — developer opens in FastReport Designer and builds the layout
await using var frxStream = await documentService.GenerateAsync(dataSet);
await frxStream.WriteToFile("MyReport.frx");

// CI mode — update Dictionary in an existing template, preserving layout
await using var templateStream = File.OpenRead("MyReport.frx");
await using var updatedStream = await documentService.GenerateAsync(dataSet, templateStream);
await updatedStream.WriteToFile("MyReport.frx");
```

---

## API Reference

### `IDocumentService`

The interface implemented by `DocumentService`, providing the contract for document generation.

### `DocumentService`

The main entry point implementing `IDocumentService`.

| Method | Return | Description |
|:--|:--|:--|
| `IsSupported(DocumentType)` | `bool` | Returns `true` for `Excel` and `Pdf` |
| `IsMergeable(DocumentType)` | `bool` | Returns `true` for `Excel` and `Pdf` |
| `MergeAsync(Stream, DataSet, DocumentType, DocumentMetadata, IDictionary?)` | `Task<Stream>` | Merges a FastReport template with data; produces PDF or delegates to `ExcelMergeAsync` |
| `ExcelMergeAsync(DataSet, DocumentMetadata, IDictionary?)` | `Task<Stream>` | Exports a `DataSet` to an `.xlsx` workbook <sup>1)</sup> |
| `GenerateAsync(DataSet, Stream?)` | `Task<Stream>` | Generates a FastReport schema document; without `templateStream` creates a new skeleton, with `templateStream` updates the Dictionary section (CI mode) <sup>2)</sup> |

<sup>1)</sup> The `parameters` argument of `ExcelMergeAsync` is part of the interface signature but is not applied during Excel export — Excel workbooks have no template parameter concept.

<sup>2)</sup> The generated stream contains UTF-8 encoded `.frx` XML. Table relations defined on the `DataSet` are included as `<Relation>` elements in the Dictionary.

### `DocumentMetadata`

Applied to generated documents. Fields and their PDF mapping:

| Property | PDF property | Description |
|:--|:--|:--|
| `Title` | Title | Document title |
| `Author` | Author | Document author |
| `Application` | Subject | Originating application name |
| `Keywords` | Keywords | Search keywords |

### Extension Methods

| Class | Method | Description |
|:--|:--|:--|
| `CellExtensions` | `SetValue(ICell, object)` | Sets an Excel cell value with correct type mapping |
| `WorkbookExtensions` | `Import(IWorkbook, DataSet)` | Imports all `DataTable`s as named worksheets; tables with no columns are skipped |
| `PDFSimpleExportExtensions` | `ApplyMetadata(PDFSimpleExport, DocumentMetadata)` | Applies `DocumentMetadata` to a `PDFSimpleExport` instance |

### Type Mapping (Excel)

| .NET Type | Excel cell type |
|:--|:--|
| `double`, `float`, `decimal`, `int`, `long`, `short`, `byte` | Numeric |
| `bool` | Boolean |
| `DateTime` | String — `yyyy-MM-dd HH:mm:ss` |
| `DateOnly` | String — `yyyy-MM-dd` |
| `null`, `DBNull` | Blank |
| All others | String (`ToString()`) |

---

## FastReport

Reports are based on [FastReport Open Source](https://github.com/FastReports/FastReport):
- Report template format: `*.frx` (XML)
- Visual designer: [FastReport Designer Community Edition](https://github.com/FastReports/FastReport/releases/latest)

---

## Build

```sh
dotnet build
dotnet pack
```

Environment variable used during build:

| Variable | Description |
|:--|:--|
| `PayrollEnginePackageDir` | Output directory for the NuGet package (optional) |

---

## Third Party Components

| Component | Purpose | License |
|:--|:--|:--|
| [NPOI](https://github.com/dotnetcore/NPOI/) | Excel export | Apache 2.0 |
| [FastReport Open Source](https://github.com/FastReports/FastReport/) | Report engine and PDF export | MIT |

---

## See Also
- [Payroll Engine WebApp](https://github.com/Payroll-Engine/PayrollEngine.WebApp) — uses this library for report rendering
- [Payroll Engine Console](https://github.com/Payroll-Engine/PayrollEngine.PayrollConsole) — uses this library for the `Report`, `DataReport` and `ReportBuild` commands
- [FastReport Open Source documentation](https://fastreports.github.io/FastReport.Documentation/) — template authoring reference
