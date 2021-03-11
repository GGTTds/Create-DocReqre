using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace CreatEXcelQWWordOT
{
    public static class Fail
    {

        public static void Aplex()
        {
            var App = new Excel.Application();
            App.SheetsInNewWorkbook = 1;
            int StartIndex = Start.StartIndex;
            Excel.Workbook workbook = App.Workbooks.Add();
            Excel.Worksheet worksheet = App.Worksheets.Item[StartIndex];
            worksheet.Name = " Накладная ";
        }
    }
}
