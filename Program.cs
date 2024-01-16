using System.Drawing;
using System.IO;
using System.Net.Mime;
using System.Text;
using Microsoft.Win32;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace lab1

{
    class Program

    {
        static void Main(string[] args)
        {
            List<string> importedStrings= ImportString();
            var strings = GetLatex(importedStrings);
            foreach (var VARIABLE in strings)
            {
                Console.WriteLine(VARIABLE);
            }
            export(strings);
        }
        public static void export(List<string> ExportStrings)
        {
            StringBuilder sb = new StringBuilder();
            foreach (var exportString in ExportStrings)
            {
                sb.Append(exportString);
            }

            System.IO.File.WriteAllText("export.txt", sb.ToString());
        }
        public static List<string> GetLatex(List<string> importedStrings)
        {

            List<string> ExportStrings = new List<string>()
            {
                @"<table style=""margin:auto""><tr><th style=""text-align:center"">Страна</th><th style=""text-align:center"">Численность населения, млн человек</th><th style=""text-align:center"">Рождаемость, %<sub>0</sub></th><th style=""text-align:center"">Смертность, %<sub>0</sub></th><th style=""text-align:center"">Доля городского населения, %</th><th style=""text-align:center"">Ожидаемая продолжительность жизни, лет</th><th style=""text-align:center"">Доля лиц в возрасте младше 15 лет, %</th><th style=""text-align:center"">Доля лиц в возрасте старше 65 лет, %</th><th style=""text-align:center"">Плотность населения, человек на 1 кв. км</th></tr></table>"
            };
            bool tmp= true;
            foreach (var elem in importedStrings)
            {
                if (elem == "Нет данных")
                {
                    ExportStrings.Add(@"<td style=""text-align:center"">" + "Нет данных" + "</td>");
                    continue;
                }
                bool result = double.TryParse(elem, out var number);
                if (result)
                {
                    ExportStrings.Add(@"<td style=""text-align:center"">" + number.ToString()+ "</td>");
                }
                else if (tmp)
                {
                    ExportStrings.Add(@"<tr><td style=""text-align:left"">" + elem + "</td>");
                    tmp=false;
                  
                }
                else
                {
                    ExportStrings.Add("</tr>");
                    ExportStrings.Add(@"<tr><td style=""text-align:left"">" + elem + "</td>");
                }
            }
            ExportStrings.Add("</tr>");
            ExportStrings.Add("</table>");
            return ExportStrings;
        }
        public static List<string> ImportString()
        {
            string fileName = "1234.xlsx";
            IWorkbook workbook;
            using (FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read))
            {
                workbook = new XSSFWorkbook(fileStream);
            }
            // Получение листа
            ISheet sheet = workbook.GetSheetAt(0);
            List<string> ExelDataInput = new List<string>();
            IRow headerRow = sheet.GetRow(0);
            for (int i = 0; i < 234; i++)
            {
                IRow row = sheet.GetRow(i);
                ICell FirstCell = row.GetCell(0);
                ExelDataInput.Add((FirstCell.ToString()));
            }
            return ExelDataInput;
        }
       
    }
}
