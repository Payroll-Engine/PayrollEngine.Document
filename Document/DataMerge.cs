using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using FastReport;
using FastReport.Export.PdfSimple;
using NPOI.XSSF.UserModel;

namespace PayrollEngine.Document;

/// <summary>Data merge, based on FastReport</summary>
public class DataMerge : IDataMerge
{
    /// <summary>Test for mergeable document</summary>
    /// <param name="documentType">The document type</param>
    /// <returns>True for Pdf documents</returns>
    public bool IsMergeable(DocumentType documentType) =>
        documentType == DocumentType.Excel ||
        documentType == DocumentType.Pdf;

    /// <inheritdoc />
    public MemoryStream Merge(Stream templateStream, DataSet dataSet, DocumentType documentType,
        DocumentMetadata metadata, IDictionary<string, object> parameters = null)
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

        // report parameters
        if (parameters != null)
        {
            foreach (var parameter in parameters)
            {
                report.SetParameterValue(parameter.Key, parameter.Value);
            }
        }

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

    /// <inheritdoc />
    public MemoryStream ExcelMerge(DataSet dataSet, DocumentMetadata metadata,
        IDictionary<string, object> parameters = null)
    {
        if (dataSet == null)
        {
            throw new ArgumentNullException(nameof(dataSet));
        }
        if (metadata == null)
        {
            throw new ArgumentNullException(nameof(metadata));
        }

        // workbook (uses the xlsx variant)
        using var workbook = new XSSFWorkbook();
        // import 
        workbook.Import(dataSet);

        // result
        var resultStream = new MemoryStream();
        workbook.Write(resultStream, true);
        resultStream.Seek(0, SeekOrigin.Begin);
        return resultStream;
    }
}