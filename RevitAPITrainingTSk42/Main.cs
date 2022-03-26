using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPITrainingTSk42
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        /*Задание 4.2 Вывод значений труб
         * Выведите в файл Excel следующие значения всех труб: имя типа трубы, наружный  диаметр трубы, внутренний диаметр трубы, длина трубы.*/
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;//Обращение к Revit
            UIDocument uidoc = uiapp.ActiveUIDocument;//Обращение к интерфейсу текущ элемента
            Document doc = uidoc.Document; //Обращение к документу


            //по классу
            var pipes = new FilteredElementCollector(doc)
                .OfClass(typeof(Pipe))
                .Cast<Pipe>()
                .ToList();

            string excelPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "pipes.xlsx");

            using (FileStream stream = new FileStream(excelPath, FileMode.Create, FileAccess.Write))
            {
                IWorkbook workBook = new XSSFWorkbook();
                ISheet sheet = workBook.CreateSheet("Лист 1");
                //  запуск метода
                int rowIndex = 0;

                double len = 0;
                double inDiam = 0;
                double outDiam = 0;

                foreach (var pipe in pipes)
                {
                    Parameter outDiamParam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER);
                    Parameter inDiamParam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM);
                    Parameter lengthParameter = pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH);
                    if (outDiamParam.StorageType == StorageType.Double)
                    {
                        outDiam = UnitUtils.ConvertFromInternalUnits(lengthParameter.AsDouble(), UnitTypeId.Millimeters);
                    } 
                    if (inDiamParam.StorageType == StorageType.Double)
                    {
                        inDiam = UnitUtils.ConvertFromInternalUnits(lengthParameter.AsDouble(), UnitTypeId.Millimeters);
                    }
                    if (lengthParameter.StorageType == StorageType.Double)
                    {
                        len = UnitUtils.ConvertFromInternalUnits(lengthParameter.AsDouble(), UnitTypeId.Meters);
                    }
                        
                    sheet.SetCellValue(rowIndex, columnIndex: 0, pipe.Name);
                    sheet.SetCellValue(rowIndex, columnIndex: 1, outDiam);
                    sheet.SetCellValue(rowIndex, columnIndex: 2, inDiam);
                    sheet.SetCellValue(rowIndex, columnIndex: 3, len);
                    rowIndex++;
                }
                workBook.Write(stream);
                workBook.Close();
            }
            //Для запуска экселя
            System.Diagnostics.Process.Start(excelPath);

            return Result.Succeeded;
        }
    }
}

