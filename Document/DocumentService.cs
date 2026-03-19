using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using FastReport;
using FastReport.Export.PdfSimple;
using NPOI.XSSF.UserModel;

namespace PayrollEngine.Document;

/// <summary>Document service, based on FastReport and NPOI</summary>
public class DocumentService : IDocumentService
{
    /// <inheritdoc />
    public bool IsSupported(DocumentType documentType) =>
        documentType == DocumentType.Excel ||
        documentType == DocumentType.Pdf;

    /// <inheritdoc />
    public bool IsMergeable(DocumentType documentType) =>
        documentType == DocumentType.Excel ||
        documentType == DocumentType.Pdf;

    /// <inheritdoc />
    public Task<Stream> MergeAsync(Stream templateStream, DataSet dataSet, DocumentType documentType,
        DocumentMetadata metadata, IDictionary<string, object> parameters = null)
    {
        ArgumentNullException.ThrowIfNull(templateStream);
        ArgumentNullException.ThrowIfNull(dataSet);
        if (!IsMergeable(documentType))
        {
            throw new ArgumentOutOfRangeException($"{documentType} merge is not supported.", nameof(documentType));
        }

        // excel
        if (documentType == DocumentType.Excel)
        {
            return ExcelMergeAsync(dataSet, metadata, parameters);
        }

        // pdf requires a FastReport template
        using var report = Report.FromStream(templateStream);
        if (parameters != null)
        {
            foreach (var parameter in parameters)
            {
                report.SetParameterValue(parameter.Key, parameter.Value);
            }
        }
        report.RegisterData(dataSet);
        report.Prepare();

        using var pdfExport = new PDFSimpleExport();
        pdfExport.ApplyMetadata(metadata);
        var resultStream = new MemoryStream();
        report.Export(pdfExport, resultStream);
        resultStream.Position = 0;

        return Task.FromResult<Stream>(resultStream);
    }

    /// <inheritdoc />
    public Task<Stream> ExcelMergeAsync(DataSet dataSet, DocumentMetadata metadata,
        IDictionary<string, object> parameters = null)
    {
        ArgumentNullException.ThrowIfNull(dataSet);
        ArgumentNullException.ThrowIfNull(metadata);

        using var workbook = new XSSFWorkbook();
        workbook.Import(dataSet);

        var resultStream = new MemoryStream();
        workbook.Write(resultStream, true);
        resultStream.Seek(0, SeekOrigin.Begin);

        return Task.FromResult<Stream>(resultStream);
    }

    /// <inheritdoc />
    public async Task<Stream> GenerateAsync(DataSet dataSet, Stream templateStream = null)
    {
        ArgumentNullException.ThrowIfNull(dataSet);

        string content;
        if (templateStream == null)
        {
            // new skeleton
            content = BuildFrxSkeleton(dataSet);
        }
        else
        {
            // CI mode: update Dictionary in existing template
            using var reader = new StreamReader(templateStream, Encoding.UTF8, leaveOpen: true);
            var existingContent = await reader.ReadToEndAsync();
            content = RebuildFrxDictionary(existingContent, dataSet);
        }

        return new MemoryStream(Encoding.UTF8.GetBytes(content));
    }

    #region Frx

    private static string BuildFrxSkeleton(DataSet dataSet)
    {
        var now = DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss", CultureInfo.InvariantCulture);
        var sb = new StringBuilder();
        sb.AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
        sb.Append("<Report ScriptLanguage=\"CSharp\"");
        sb.Append($" ReportInfo.Created=\"{now}\"");
        sb.Append($" ReportInfo.Modified=\"{now}\"");
        sb.AppendLine(" ReportInfo.CreatorVersion=\"2023.2.0.0\">");
        sb.AppendLine(BuildFrxDictionary(dataSet));
        sb.AppendLine("  <ReportPage Name=\"Page1\" Watermark.Font=\"Arial, 60pt\">");
        sb.AppendLine("  </ReportPage>");
        sb.Append("</Report>");
        return sb.ToString();
    }

    private static string RebuildFrxDictionary(string frxContent, DataSet dataSet)
    {
        var doc = new XmlDocument();
        doc.LoadXml(frxContent);

        var root = doc.DocumentElement
            ?? throw new InvalidOperationException("Invalid .frx file: missing root element.");

        // remove existing Dictionary
        var existing = root.SelectSingleNode("Dictionary");
        if (existing != null)
        {
            root.RemoveChild(existing);
        }

        // import new Dictionary
        var dictXml = new XmlDocument();
        dictXml.LoadXml($"<root>{BuildFrxDictionary(dataSet)}</root>");
        var dictNode = dictXml.DocumentElement!.SelectSingleNode("Dictionary");
        var imported = doc.ImportNode(dictNode!, true);
        root.InsertBefore(imported, root.FirstChild);

        // serialize with indentation
        var sb = new StringBuilder();
        var settings = new XmlWriterSettings
        {
            Indent = true,
            IndentChars = "  ",
            Encoding = Encoding.UTF8,
            OmitXmlDeclaration = false
        };
        using var writer = XmlWriter.Create(sb, settings);
        doc.Save(writer);
        return sb.ToString();
    }

    private static string BuildFrxDictionary(DataSet dataSet)
    {
        var sb = new StringBuilder();
        sb.AppendLine("  <Dictionary>");

        // tables
        foreach (DataTable table in dataSet.Tables)
        {
            sb.AppendLine($"    <TableDataSource Name=\"{table.TableName}\" ReferenceName=\"Data.{table.TableName}\" DataType=\"System.Int32\" Enabled=\"true\">");
            foreach (DataColumn column in table.Columns)
            {
                sb.AppendLine($"      <Column Name=\"{column.ColumnName}\" DataType=\"{column.DataType.FullName}\"/>");
            }
            sb.AppendLine("    </TableDataSource>");
        }

        // relations
        foreach (DataRelation relation in dataSet.Relations)
        {
            var parentColumns = string.Join(",", Array.ConvertAll(relation.ParentColumns, c => c.ColumnName));
            var childColumns = string.Join(",", Array.ConvertAll(relation.ChildColumns, c => c.ColumnName));
            sb.AppendLine($"    <Relation Name=\"{relation.RelationName}\"" +
                          $" ParentDataSource=\"{relation.ParentTable.TableName}\"" +
                          $" ChildDataSource=\"{relation.ChildTable.TableName}\"" +
                          $" ParentColumns=\"{parentColumns}\"" +
                          $" ChildColumns=\"{childColumns}\"" +
                          $" Enabled=\"true\"/>");
        }

        sb.Append("  </Dictionary>");
        return sb.ToString();
    }

    #endregion
}
