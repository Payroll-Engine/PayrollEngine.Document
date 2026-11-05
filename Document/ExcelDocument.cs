using System;
using System.Data;
using System.Globalization;
using System.IO;
using ClosedXML.Excel;

namespace PayrollEngine.Document;

/// <summary>Excel document writer</summary>
internal static class ExcelDocument
{
    /// <summary>Maximum length of a worksheet name</summary>
    private const int MaxSheetNameLength = 31;

    /// <summary>Write a data set as Excel document (one worksheet per table)</summary>
    /// <remarks>The header row contains the column names and is used as filter row.
    /// Column widths are adjusted to the content.</remarks>
    /// <param name="dataSet">The data set to write</param>
    /// <param name="stream">The target stream</param>
    internal static void Write(DataSet dataSet, Stream stream)
    {
        ArgumentNullException.ThrowIfNull(dataSet);
        ArgumentNullException.ThrowIfNull(stream);

        using var workbook = new XLWorkbook();
        foreach (DataTable table in dataSet.Tables)
        {
            AddWorksheet(workbook, table);
        }
        // a workbook without worksheet cannot be saved
        if (workbook.Worksheets.Count == 0)
        {
            workbook.Worksheets.Add(GetSheetName(dataSet.DataSetName));
        }
        workbook.SaveAs(stream);
    }

    private static void AddWorksheet(XLWorkbook workbook, DataTable table)
    {
        // empty table
        if (table.Columns.Count == 0)
        {
            return;
        }

        var sheet = workbook.Worksheets.Add(GetSheetName(table.TableName));

        // header with the table column names
        for (var x = 0; x < table.Columns.Count; x++)
        {
            sheet.Cell(1, x + 1).Value = table.Columns[x].ColumnName;
        }

        // use header as filter row
        sheet.Range(1, 1, 1, table.Columns.Count).SetAutoFilter();

        // data rows
        for (var y = 0; y < table.Rows.Count; y++)
        {
            var tableRow = table.Rows[y];
            for (var x = 0; x < table.Columns.Count; x++)
            {
                sheet.Cell(y + 2, x + 1).Value = ToCellValue(tableRow[x]);
            }
        }

        // auto size column widths
        sheet.Columns(1, table.Columns.Count).AdjustToContents();
    }

    private static string GetSheetName(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return "Sheet1";
        }
        return name.Length > MaxSheetNameLength ? name[..MaxSheetNameLength] : name;
    }

    private static XLCellValue ToCellValue(object value)
    {
        switch (value)
        {
            case null or DBNull:
                return Blank.Value;
            case double doubleValue:
                return doubleValue;
            case float floatValue:
                return floatValue;
            case decimal decimalValue:
                return (double)decimalValue;
            case int intValue:
                return intValue;
            case long longValue:
                return longValue;
            case short shortValue:
                return shortValue;
            case byte byteValue:
                return byteValue;
            case bool boolValue:
                return boolValue;
            // dates are written as text
            case DateOnly dateOnly:
                return dateOnly.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            case DateTime dateTime:
                return dateTime.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            default:
                return value.ToString();
        }
    }
}
