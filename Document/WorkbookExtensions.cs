using System;
using System.Data;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace PayrollEngine.Document;

/// <summary>Extensions for the type <see cref="IWorkbook"/> </summary>
public static class WorkbookExtensions
{
    /// <summary>Import data set to a workbook</summary>
    /// <param name="workbook">The target workbook</param>
    /// <param name="dataSet">The data set to import</param>
    public static void Import(this IWorkbook workbook, DataSet dataSet)
    {
        if (dataSet == null)
        {
            throw new ArgumentNullException(nameof(dataSet));
        }

        // for any data table a worksheet
        foreach (DataTable table in dataSet.Tables)
        {
            // empty table
            if (table.Columns.Count == 0)
            {
                continue;
            }

            // sheet
            var sheet = workbook.CreateSheet(table.TableName);

            // header with the table column names
            var headerRow = sheet.CreateRow(0);
            for (var x = 0; x < table.Columns.Count; x++)
            {
                var cell = headerRow.CreateCell(x);
                cell.SetCellValue(table.Columns[x].ColumnName);
            }

            // use header as filter row
            sheet.SetAutoFilter(new CellRangeAddress(0, 0, 0, table.Columns.Count - 1));

            // data rows
            for (var y = 0; y < table.Rows.Count; y++)
            {
                var row = sheet.CreateRow(y + 1);
                var tableRow = table.Rows[y];
                // column value
                for (var x = 0; x < table.Columns.Count; x++)
                {
                    var cell = row.CreateCell(x);
                    var dataValue = tableRow[x];
                    cell.SetValue(dataValue);
                }
            }

            // auto size column widths
            foreach (DataColumn column in table.Columns)
            {
                sheet.AutoSizeColumn(column.Ordinal);
            }
        }
    }
}