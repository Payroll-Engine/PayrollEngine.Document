using System;
using System.Globalization;
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
        if (value == null || value is DBNull)
        {
            cell.SetBlank();
            return;
        }

        switch (value)
        {
            case double doubleValue:
                cell.SetCellValue(doubleValue);
                break;
            case float floatValue:
                cell.SetCellValue(floatValue);
                break;
            case decimal decimalValue:
                cell.SetCellValue((double)decimalValue);
                break;
            case int intValue:
                cell.SetCellValue(intValue);
                break;
            case long longValue:
                cell.SetCellValue(longValue);
                break;
            case short shortValue:
                cell.SetCellValue(shortValue);
                break;
            case byte byteValue:
                cell.SetCellValue(byteValue);
                break;
            case bool boolValue:
                cell.SetCellValue(boolValue);
                break;
            case DateOnly dateOnly:
                cell.SetCellValue(dateOnly.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
                break;
            case DateTime dateTime:
                cell.SetCellValue(dateTime.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture));
                break;
            default:
                cell.SetCellValue(value.ToString());
                break;
        }
    }
}