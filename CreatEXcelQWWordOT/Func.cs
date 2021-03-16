using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows;
using System.Diagnostics;



namespace CreatEXcelQWWordOT
{

    public class Func
    {
        public static int ddd;

            public static void BP114F()
            {
                StreamReader EDD = new StreamReader("log1.txt");
                string lim;
                lim = EDD.ReadLine().ToString();




                if (Start.ВР114F == true)
                {

                    //Excel.Range RR1 = worksheet2.Range[worksheet2.Cells[1][StartIndex], worksheet2.Cells[6][StartIndex + 5]];
                    //RR1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                    //    RR1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                    //    RR1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                    //    RR1.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                    //    RR1.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                    //    RR1.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                    //worksheet2.Cells[1][StartIndex] = worksheet1.Cells[1][StartIndex];
                    //worksheet2.Cells[2][StartIndex] = worksheet1.Cells[2][StartIndex];
                    //worksheet2.Cells[6][StartIndex] = worksheet1.Cells[6][StartIndex];
                    //Excel.Range Head21 = worksheet2.Range[worksheet2.Cells[2][StartIndex], worksheet2.Cells[5][StartIndex]];
                    //Head21.Merge();
                    //StartIndex +=1;
                    //worksheet2.Cells[1][StartIndex] = worksheet1.Cells[1][StartIndex];
                    //worksheet2.Cells[2][StartIndex] = worksheet1.Cells[2][StartIndex];
                    //Excel.Range Head211 = worksheet2.Range[worksheet2.Cells[2][StartIndex], worksheet2.Cells[5][StartIndex]];
                    //Head211.Merge();
                    //StartIndex +=1;
                    //worksheet2.Cells[1][StartIndex] = worksheet1.Cells[1][StartIndex];
                    //worksheet2.Cells[2][StartIndex] = worksheet1.Cells[2][StartIndex];
                    //Excel.Range Head214 = worksheet2.Range[worksheet2.Cells[2][StartIndex], worksheet2.Cells[5][StartIndex]];
                    //Head214.Merge();
                    //StartIndex +=1;
                    //worksheet2.Cells[1][StartIndex] = worksheet1.Cells[1][StartIndex];
                    //worksheet2.Cells[2][StartIndex] = worksheet1.Cells[2][StartIndex];
                    //worksheet2.Cells[3][StartIndex] = worksheet1.Cells[3][StartIndex];
                    //worksheet2.Cells[4][StartIndex] = worksheet1.Cells[4][StartIndex];
                    //worksheet2.Cells[5][StartIndex] = worksheet1.Cells[5][StartIndex];
                    //StartIndex +=1;
                    //worksheet2.Cells[1][StartIndex] = worksheet1.Cells[1][StartIndex];
                    //worksheet2.Cells[2][StartIndex] = worksheet1.Cells[2][StartIndex];
                    //StartIndex += 1;
                    //worksheet2.Cells[1][StartIndex] = worksheet1.Cells[1][StartIndex];
                    //worksheet2.Cells[2][StartIndex] = worksheet1.Cells[2][StartIndex];
                    //Excel.Range Head2 = worksheet2.Range[worksheet2.Cells[6][2], worksheet2.Cells[6][StartIndex]];
                    //Head2.Merge();
                    //StartIndex += 1;
                    //string g = "Лист2";
                    //Viz(g);

                }
                else { }

                if (Start.ВР2F == true)

                {
                    //string g = "Лист3";
                    //Viz(g);
                
                }

                //  }
                //  else { }
                //  if (Start.ZN34 == true)
                //  {
                //      Excel.Range RR1 = worksheet.Range[worksheet.Cells[1][StartIndex], worksheet.Cells[5][StartIndex + 9]];
                //      RR1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                //          RR1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                //          RR1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                //          RR1.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                //          RR1.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                //          RR1.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;

                //      worksheet.Cells[1][StartIndex] = " Завод ";
                //      StartIndex += 1;
                //      worksheet.Cells[1][StartIndex] = " Наименование ";
                //      StartIndex += 1;
                //      worksheet.Cells[1][StartIndex] = " Поступило штук ";
                //      StartIndex += 1;
                //      worksheet.Cells[1][StartIndex] = " П№ чертежа ";
                //      StartIndex += 1;
                //      worksheet.Cells[1][StartIndex] = " ВидПроверки ";
                //      worksheet.Cells[2][StartIndex] = " Норма ";
                //      worksheet.Cells[3][StartIndex] = " Факт ";
                //      worksheet.Cells[4][StartIndex] = " Проверено, шт ";
                //      worksheet.Cells[5][StartIndex] = " Несоотв., шт ";
                //      StartIndex += 1;
                //      worksheet.Cells[1][StartIndex] = "Ø посадочного отверстия, мм ";
                //      StartIndex += 1;
                //      worksheet.Cells[1][StartIndex] = " Ø бурта наружный, мм ";
                //      StartIndex += 1;
                //      worksheet.Cells[1][StartIndex] = " Высота, мм ";
                //      StartIndex += 1;
                //      worksheet.Cells[1][StartIndex] = " Ø наружный, мм ";
                //      StartIndex += 1;
                //      worksheet.Cells[1][StartIndex] = " Масса,г ";


                //      StartIndex -= 9;

                //      worksheet.Cells[2][StartIndex] = Start.Zav;
                //      StartIndex += 1;
                //      worksheet.Cells[2][StartIndex] = DataElemProduc.ZN34.Ima.ToString(); ;
                //      StartIndex += 1;
                //      worksheet.Cells[2][StartIndex] = Start.KolTov;
                //      StartIndex += 1;
                //      worksheet.Cells[2][StartIndex] = DataElemProduc.ZN34.Chert;
                //      StartIndex += 1;
                //      StartIndex += 1;
                //      worksheet.Cells[2][StartIndex] = $"{DataElemProduc.ZN34.PosOtv} - {DataElemProduc.ZN34.Pog}";
                //      StartIndex += 1;
                //      worksheet.Cells[2][StartIndex] = DataElemProduc.ZN34.BurtNar;
                //      StartIndex += 1;
                //      worksheet.Cells[2][StartIndex] = DataElemProduc.ZN34.Hei;
                //      StartIndex += 1;
                //      worksheet.Cells[2][StartIndex] = DataElemProduc.ZN34.NarDia;
                //      StartIndex += 1;
                //      worksheet.Cells[2][StartIndex] = " Напишите массу";





                //      worksheet.Columns.AutoFit();


                //  }




            }


            public static void In6Rows()
            {

            }

        public static void Viz(string[] s)
        {

            var App = new Excel.Application();
            Excel.Workbook xlWB;

            string xlFileName = "E:/Практика от 01.03.21/CreatEXcelQWWordOT/CreatEXcelQWWordOT/bin/Debug/Form.xlsx";
            xlWB = App.Workbooks.Open(xlFileName.ToString()); //открываем наш файл 
            int StartIndex = 1;
            //MessageBox.Show(s[1].ToString());


            for (int i=0; i<=Start.KolPov; i++)
            {
              
                try
                {


                    Excel.Range Rng;
                 

                    Excel.Worksheet worksheet1 = App.Worksheets[s[i]];
                    Excel.Worksheet worksheet2 = App.Worksheets["Отчет"];
                    Excel.Range RR1 = worksheet2.Range[worksheet2.Cells[1][StartIndex], worksheet2.Cells[6][StartIndex + 5]];
                    RR1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                        RR1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                        RR1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                        RR1.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                        RR1.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                        RR1.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                    worksheet2.Cells[1][StartIndex] = worksheet1.Cells[1][1];
                    worksheet2.Cells[2][StartIndex] = worksheet1.Cells[2][1];
                    worksheet2.Cells[6][StartIndex] = worksheet1.Cells[6][1];
                    Excel.Range Head21 = worksheet2.Range[worksheet2.Cells[2][StartIndex], worksheet2.Cells[5][StartIndex]];
                    Head21.Merge();
                    StartIndex += 1;
                    worksheet2.Cells[1][StartIndex] = worksheet1.Cells[1][2];
                    worksheet2.Cells[2][StartIndex] = worksheet1.Cells[2][2];
                    Excel.Range Head211 = worksheet2.Range[worksheet2.Cells[2][StartIndex], worksheet2.Cells[5][StartIndex]];
                    Head211.Merge();
                    StartIndex += 1;
                    worksheet2.Cells[1][StartIndex] = worksheet1.Cells[1][3];
                    worksheet2.Cells[2][StartIndex] = worksheet1.Cells[2][3];
                    Excel.Range Head214 = worksheet2.Range[worksheet2.Cells[2][StartIndex], worksheet2.Cells[5][StartIndex]];
                    Head214.Merge();
                    StartIndex += 1;
                    worksheet2.Cells[1][StartIndex] = worksheet1.Cells[1][4];
                    worksheet2.Cells[2][StartIndex] = worksheet1.Cells[2][4];
                    worksheet2.Cells[3][StartIndex] = worksheet1.Cells[3][4];
                    worksheet2.Cells[4][StartIndex] = worksheet1.Cells[4][4];
                    worksheet2.Cells[5][StartIndex] = worksheet1.Cells[5][4];
                    StartIndex += 1;
                    worksheet2.Cells[1][StartIndex] = worksheet1.Cells[1][5];
                    worksheet2.Cells[2][StartIndex] = worksheet1.Cells[2][5];
                    StartIndex += 1;
                    worksheet2.Cells[1][StartIndex] = worksheet1.Cells[1][6];
                    worksheet2.Cells[2][StartIndex] = worksheet1.Cells[2][6];
                    //Excel.Range Head2 = worksheet2.Range[worksheet2.Cells[6][2], worksheet2.Cells[6][StartIndex]];
                    //Head2.Merge();
                    StartIndex += 1;
                    if (worksheet2.Cells[1][StartIndex] != null)
                    {
                        worksheet2.Cells[1][StartIndex] = worksheet1.Cells[1][7];
                        worksheet2.Cells[2][StartIndex] = worksheet1.Cells[2][7];
                    }
                    StartIndex += 1;
                    if (worksheet2.Cells[1][StartIndex] != null)
                    {
                        worksheet2.Cells[1][StartIndex] = worksheet1.Cells[1][8];
                        worksheet2.Cells[2][StartIndex] = worksheet1.Cells[2][8];
                    }
                    worksheet2.Columns.AutoFit();
                }

                catch
                {

                }


            }
            
            App.Visible = true;
    


        }



        }
    }
