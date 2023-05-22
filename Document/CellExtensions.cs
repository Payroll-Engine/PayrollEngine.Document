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
                cell.SetCellType(CellType.Numeric);
                break;
            case bool boolValue:
                cell.SetCellValue(boolValue);
                cell.SetCellType(CellType.Boolean);
                break;
            case DateOnly dateOnly:
                cell.SetCellValue(dateOnly.ToShortDateString());
                cell.SetCellType(CellType.String);
                break;
            case DateTime dateTime:
                cell.SetCellValue(
                    $"{dateTime.ToShortDateString()} {dateTime.ToShortTimeString()}");
                cell.SetCellType(CellType.String);
                break;
            default:
                cell.SetCellValue(value.ToString());
                cell.SetCellType(CellType.String);
                break;
        }
    }
}