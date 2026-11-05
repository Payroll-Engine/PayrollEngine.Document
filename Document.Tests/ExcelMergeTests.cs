using System.Data;
using ClosedXML.Excel;
using Xunit;

namespace PayrollEngine.Document.Tests;

/// <summary>Excel output of <see cref="DocumentService.ExcelMergeAsync"/>, read back with ClosedXML</summary>
public class ExcelMergeTests
{
    private static async Task<XLWorkbook> MergeAsync(DataSet dataSet)
    {
        var stream = await new DocumentService().ExcelMergeAsync(dataSet, new DocumentMetadata());
        return new XLWorkbook(stream);
    }

    private static DataTable CreateTable(string name)
    {
        var table = new DataTable(name);
        table.Columns.Add("Text", typeof(string));
        table.Columns.Add("Decimal", typeof(decimal));
        table.Columns.Add("Integer", typeof(int));
        table.Columns.Add("Boolean", typeof(bool));
        table.Columns.Add("DateTime", typeof(DateTime));
        table.Columns.Add("Date", typeof(DateOnly));
        table.Columns.Add("Empty", typeof(string));
        table.Rows.Add("Alpha", 1234.56m, 42, true,
            new DateTime(2026, 1, 31, 8, 5, 0), new DateOnly(2026, 2, 28), DBNull.Value);
        return table;
    }

    [Fact]
    public async Task OneWorksheetPerTable()
    {
        var dataSet = new DataSet("Report");
        dataSet.Tables.Add(CreateTable("Employees"));
        dataSet.Tables.Add(CreateTable("Wages"));

        using var workbook = await MergeAsync(dataSet);

        Assert.Equal(["Employees", "Wages"], workbook.Worksheets.Select(x => x.Name));
    }

    [Fact]
    public async Task HeaderRowIsFilterRow()
    {
        var dataSet = new DataSet("Report");
        dataSet.Tables.Add(CreateTable("Employees"));

        using var workbook = await MergeAsync(dataSet);
        var sheet = workbook.Worksheet(1);

        Assert.Equal(["Text", "Decimal", "Integer", "Boolean", "DateTime", "Date", "Empty"],
            Enumerable.Range(1, 7).Select(x => sheet.Cell(1, x).GetString()));
        Assert.Equal("A1:G1", sheet.AutoFilter.Range.RangeAddress.ToString());
    }

    [Fact]
    public async Task CellValues()
    {
        var dataSet = new DataSet("Report");
        dataSet.Tables.Add(CreateTable("Employees"));

        using var workbook = await MergeAsync(dataSet);
        var sheet = workbook.Worksheet(1);

        Assert.Equal("Alpha", sheet.Cell(2, 1).GetText());
        Assert.Equal(1234.56, sheet.Cell(2, 2).GetDouble());
        Assert.Equal(42, sheet.Cell(2, 3).GetDouble());
        Assert.True(sheet.Cell(2, 4).GetBoolean());
        // dates are written as text
        Assert.Equal("2026-01-31 08:05:00", sheet.Cell(2, 5).GetText());
        Assert.Equal("2026-02-28", sheet.Cell(2, 6).GetText());
        Assert.True(sheet.Cell(2, 7).Value.IsBlank);
    }

    [Fact]
    public async Task ColumnWidthsAdjustedToContent()
    {
        var table = new DataTable("Texts");
        table.Columns.Add("A", typeof(string));
        table.Columns.Add("B", typeof(string));
        table.Rows.Add("x", new string('w', 60));
        var dataSet = new DataSet("Report");
        dataSet.Tables.Add(table);

        using var workbook = await MergeAsync(dataSet);
        var sheet = workbook.Worksheet(1);

        Assert.True(sheet.Column(2).Width > sheet.Column(1).Width * 3);
    }

    [Fact]
    public async Task LongTableNameTruncatedToSheetNameLimit()
    {
        var dataSet = new DataSet("Report");
        dataSet.Tables.Add(CreateTable(new string('T', 40)));

        using var workbook = await MergeAsync(dataSet);

        Assert.Equal(new string('T', 31), workbook.Worksheet(1).Name);
    }

    [Fact]
    public async Task TableWithoutColumnsIsSkipped()
    {
        var dataSet = new DataSet("Report");
        dataSet.Tables.Add(new DataTable("NoColumns"));
        dataSet.Tables.Add(CreateTable("Employees"));

        using var workbook = await MergeAsync(dataSet);

        Assert.Equal(["Employees"], workbook.Worksheets.Select(x => x.Name));
    }

    [Fact]
    public async Task DataSetWithoutColumnsGivesEmptyWorksheet()
    {
        var dataSet = new DataSet("Report");
        dataSet.Tables.Add(new DataTable("NoColumns"));

        using var workbook = await MergeAsync(dataSet);

        var sheet = Assert.Single(workbook.Worksheets);
        Assert.Equal("Report", sheet.Name);
        Assert.Null(sheet.LastCellUsed());
    }

    [Fact]
    public async Task MissingDataSetThrows() =>
        await Assert.ThrowsAsync<ArgumentNullException>(() =>
            new DocumentService().ExcelMergeAsync(null, new DocumentMetadata()));
}
