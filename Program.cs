using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace chart_to_picture
{
    class Program
    {
        const string EXPORT_TO_DIRECTORY = @"C:\Users\etanik\Desktop\images\";
        static void Main(string[] excel)
        {
            Excel.Application imgconverter = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application") as Microsoft.Office.Interop.Excel.Application;

            ConsoleColor c = Console.ForegroundColor;
            Console.ForegroundColor = ConsoleColor.Red;
            Console.Write("Export To: ");
            Console.ForegroundColor = c;
            string exportPath = Console.ReadLine();

            if (exportPath == "")
                exportPath = EXPORT_TO_DIRECTORY;
            Excel.Workbook wb = imgconverter.ActiveWorkbook;

            foreach(Excel.Worksheet ws in wb.Worksheets)

            {
                Excel.ChartObjects chartobjects = (Excel.ChartObjects)(ws.ChartObjects(Type.Missing));

                foreach(Excel.ChartObject co in chartobjects)

                {
                    co.Select();
                    Excel.Chart chart = co.Chart;
                    chart.Export(exportPath + chart.Name + ".png", "PNG", false);
                }
            }

            Process.Start(exportPath);

        }
    }
}
