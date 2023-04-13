using FastReport.Export.PdfSimple;

namespace PayrollEngine.Document;

/// <summary>Extensions for the type <see cref="PDFSimpleExport"/> </summary>
public static class PDFSimpleExportExtensions
{
    /// <summary>Apply metadata to the document</summary>
    /// <param name="export">The document</param>
    /// <param name="metadata">The document meta data</param>
    public static void ApplyMetadata(this PDFSimpleExport export, DocumentMetadata metadata)
    {
        if (export == null || metadata == null)
        {
            return;
        }

        // built-in properties
        if (!string.IsNullOrWhiteSpace(metadata.Application))
        {
            export.Subject = metadata.Application;
        }
        if (!string.IsNullOrWhiteSpace(metadata.Author))
        {
            export.Author = metadata.Author;
        }
        //  document.DocumentInformation.Company = metadata.Company;
        if (!string.IsNullOrWhiteSpace(metadata.Title))
        {
            export.Title = metadata.Title;
        }
        //document.DocumentInformation.Category = metadata.Category;
        if (!string.IsNullOrWhiteSpace(metadata.Keywords))
        {
            export.Keywords = metadata.Keywords;
        }
    }
}