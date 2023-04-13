using System;
using System.Data;
using NPOI.HSSF.UserModel;

namespace PayrollEngine.Document;

/// <summary>Extensions for the type <see cref="HSSFWorkbook"/> </summary>
public static class WorkbookExtensions
{
    /// <summary>Import data set to a workbook</summary>
    /// <param name="workbook">The target workbook</param>
    /// <param name="dataSet">The data set to import</param>
    public static void Import(this HSSFWorkbook workbook, DataSet dataSet)
    {
        if (dataSet == null)
        {
            throw new ArgumentNullException(nameof(dataSet));
        }

        // for any data table a worksheet
        foreach (DataTable table in dataSet.Tables)
        {
            var sheet = workbook.CreateSheet(table.TableName);

            // header with the table column names
            var headerRow = sheet.CreateRow(0);
            for (var x = 0; x < table.Columns.Count; x++)
            {
                var cell = headerRow.CreateCell(x);
                cell.SetCellValue(table.Columns[x].ColumnName);
            }

            // data rows
            for (var y = 0; y < table.Rows.Count; y++)
            {
                var row = sheet.CreateRow(y + 1);
                var tableRow = table.Rows[y];
                for (var x = 0; x < table.Columns.Count; x++)
                {
                    var cell = row.CreateCell(x);
                    var dataValue = tableRow[x];
                    cell.SetValue(dataValue);
                }
            }
        }
    }
}