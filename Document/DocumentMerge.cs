using System;
using System.Data;
using System.IO;
using FastReport;
using FastReport.Export.PdfSimple;
using NPOI.HSSF.UserModel;

namespace PayrollEngine.Document;

/// <summary>Document merge, based on FastReport</summary>
public class DocumentMerge : IDocumentMerge
{
    /// <summary>Test for mergeable document</summary>
    /// <param name="documentType">The document type</param>
    /// <returns>True for Pdf documents</returns>
    public bool IsMergeable(DocumentType documentType) =>
        documentType == DocumentType.Pdf;

    /// <summary>
    /// Merge a document from a stream to a stream
    /// </summary>
    /// <param name="templateStream">Name of the template file</param>
    /// <param name="dataSet">The data set</param>
    /// <param name="documentType">Type of the document</param>
    /// <param name="metadata">The document metadata</param>
    /// <returns>The merged document stream</returns>
    public MemoryStream Merge(Stream templateStream, DataSet dataSet, DocumentType documentType,
        DocumentMetadata metadata)
    {
        if (templateStream == null)
        {
            throw new ArgumentNullException(nameof(templateStream));
        }
        if (dataSet == null)
        {
            throw new ArgumentNullException(nameof(dataSet));
        }
        if (!IsMergeable(documentType))
        {
            throw new ArgumentOutOfRangeException($"{documentType} merge is not supported", nameof(documentType));
        }

        // excel
        if (documentType == DocumentType.Excel)
        {
            return ExcelMerge(dataSet, metadata);
        }

        // word and pdf requires a word template
        using var report = Report.FromStream(templateStream);

        // register the data set
        report.RegisterData(dataSet);

        // prepare the report
        report.Prepare();

        // export to pdf
        var pdfExport = new PDFSimpleExport();
        pdfExport.ApplyMetadata(metadata);
        var resultStream = new MemoryStream();
        report.Export(pdfExport, resultStream);

        resultStream.Position = 0;
        return resultStream;
    }

    /// <summary>Merge to excel stream</summary>
    /// <param name="dataSet">The source data</param>
    /// <param name="metadata">The document meta data</param>
    public MemoryStream ExcelMerge(DataSet dataSet, DocumentMetadata metadata)
    {
        if (dataSet == null)
        {
            throw new ArgumentNullException(nameof(dataSet));
        }
        if (metadata == null)
        {
            throw new ArgumentNullException(nameof(metadata));
        }

        // workbook
        using var workbook = new HSSFWorkbook();
        // import 
        workbook.Import(dataSet);

        // result
        var resultStream = new MemoryStream();
        workbook.Write(resultStream);
        resultStream.Position = 0;
        return resultStream;
    }
}