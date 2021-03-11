using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace CreatEXcelQWWordOT
{
   public  class Func
    {
        public static int ddd;
        public static void BP114F()
        {
            var App = new Excel.Application();
            App.SheetsInNewWorkbook = 1;
            int StartIndex = Start.StartIndex;
            Excel.Workbook workbook = App.Workbooks.Add();
            Excel.Worksheet worksheet = App.Worksheets.Item[StartIndex];
            worksheet.Name = " Накладная ";

            worksheet.Cells[1][StartIndex] = " Завод ";
            StartIndex += 1;
            worksheet.Cells[1][StartIndex ] = " Наименование ";
            StartIndex += 1;
            worksheet.Cells[1][StartIndex ] = " Поступило штук ";
            StartIndex += 1;
            worksheet.Cells[1][StartIndex ] = " П№ чертежа ";
            StartIndex += 1;
            worksheet.Cells[1][StartIndex ] = " ВидПроверки ";
            worksheet.Cells[2][StartIndex ] = " Норма ";
            worksheet.Cells[3][StartIndex ] = " Факт ";
            worksheet.Cells[4][StartIndex ] = " Проверено, шт ";
            worksheet.Cells[5][StartIndex ] = " Несоотв., шт ";
            StartIndex += 1;
            worksheet.Cells[1][StartIndex] = " Посадочного отв., мм ";
            StartIndex += 1;
            worksheet.Cells[1][StartIndex] = " Масса,г ";
            

            StartIndex -=6;

            worksheet.Cells[2][StartIndex] = Start.Zav;
            StartIndex += 1;
            worksheet.Cells[2][StartIndex] = DataElemProduc.ВР114F.Ima.ToString(); ;
            StartIndex += 1;
            worksheet.Cells[2][StartIndex] = Start.KolTov;
            StartIndex += 1;
            worksheet.Cells[2][StartIndex] = DataElemProduc.ВР114F.Chert;
            StartIndex += 1;
            StartIndex += 1;
            worksheet.Cells[2][StartIndex] = $"{DataElemProduc.ВР114F.PosOtv} - {DataElemProduc.ВР114F.Pog}";
            StartIndex += 1;
            worksheet.Cells[2][StartIndex] = " Напишите массу";
           
            Excel.Range RR = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[5][StartIndex]];
            RR.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                RR.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                RR.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                RR.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                RR.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                RR.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;



            worksheet.Columns.AutoFit();

            App.Visible = true;
        }





        public static void ВР2F()
        {
            var App = new Excel.Application();
                App.SheetsInNewWorkbook = 1;
                int StartIndex = Start.StartIndex;
                Excel.Workbook workbook = App.Workbooks.Add();
                Excel.Worksheet worksheet = App.Worksheets.Item[StartIndex];
                worksheet.Name = " Накладная ";
           

            worksheet.Cells[1][StartIndex] = " Завод ";
            StartIndex += 1;
            worksheet.Cells[1][StartIndex] = " Наименование ";
            StartIndex += 1;
            worksheet.Cells[1][StartIndex] = " Поступило штук ";
            StartIndex += 1;
            worksheet.Cells[1][StartIndex] = " П№ чертежа ";
            StartIndex += 1;
            worksheet.Cells[1][StartIndex] = " ВидПроверки ";
            worksheet.Cells[2][StartIndex] = " Норма ";
            worksheet.Cells[3][StartIndex] = " Факт ";
            worksheet.Cells[4][StartIndex] = " Проверено, шт ";
            worksheet.Cells[5][StartIndex] = " Несоотв., шт ";
            StartIndex += 1;
            worksheet.Cells[1][StartIndex] = " Посадочного отв., мм ";
            StartIndex += 1;
            worksheet.Cells[1][StartIndex] = " Масса,г ";


            StartIndex -= 6;

            worksheet.Cells[2][StartIndex] = Start.Zav;
            StartIndex += 1;
            worksheet.Cells[2][StartIndex] = DataElemProduc.ВР2F.Ima.ToString(); ;
            StartIndex += 1;
            worksheet.Cells[2][StartIndex] = Start.KolTov;
            StartIndex += 1;
            worksheet.Cells[2][StartIndex] = DataElemProduc.ВР2F.Chert;
            StartIndex += 1;
            StartIndex += 1;
            worksheet.Cells[2][StartIndex] = $"{DataElemProduc.ВР2F.PosOtv} - {DataElemProduc.ВР2F.Pog}";
            StartIndex += 1;
            worksheet.Cells[2][StartIndex] = " Напишите массу";

            Excel.Range RR = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[5][StartIndex]];
            RR.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                RR.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                RR.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                RR.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                RR.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                RR.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;



            worksheet.Columns.AutoFit();

            App.Visible = true;
        }



        public static void In6Rows()
        {
         
        }
    
    
    
    
    
    
    }
}
