using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPITrainingTSk42
{
    //метод для записи значений в ячейки
    public static class SheetExts
    {//обобщенный метод для записи
        public static void SetCellValue<T>(this ISheet sheet, int rowIndex, int columnIndex, T value)
        {
            var cellReference = new CellReference(rowIndex, columnIndex);
            //создание и проверка строки
            var row = sheet.GetRow(cellReference.Row);
            if (row == null)
                row = sheet.CreateRow(cellReference.Row);
            //создание ячейки
            var cell = row.GetCell(cellReference.Col);
            if (cell == null)
                cell = row.CreateCell(cellReference.Col);
            //запсиь в ячейку
            if (value is string)
            {
                cell.SetCellValue((string)(object)value);
            }
            else if (value is double)
            {
                cell.SetCellValue((double)(object)value);
            }
            else if (value is int)
            {
                cell.SetCellValue((int)(object)value);
            }
        }
    }
}