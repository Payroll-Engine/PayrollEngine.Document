using System;
using NPOI.SS.UserModel;

namespace PayrollEngine.Document;

/// <summary>Extensions for the type <see cref="ICell"/> </summary>
public static class CellExtensions
{

    /// <summary>Set the excel cell value</summary>
    /// <param name="cell">The target cell</param>
    /// <param name="value">The value to set</param>
    public static void SetValue(this ICell cell, object value)
    {
        switch (value)
        {
            case double doubleValue:
                cell.SetCellValue(doubleValue);
                break;
            case bool boolValue:
                cell.SetCellValue(boolValue);
                break;
            case DateOnly dateOnly:
                cell.SetCellValue(dateOnly);
                break;
            case DateTime dateTime:
                cell.SetCellValue(dateTime);
                break;
            default:
                cell.SetCellValue(value.ToString());
                break;
        }
    }
}