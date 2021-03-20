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


        public static void Viz(string[] s)
        {

            var App = new Excel.Application();
            Excel.Workbook xlWB;

            string L;
            StreamReader rr = new StreamReader("Put.txt");
            L = rr.ReadLine();
            //"E:/.../CreatEXcelQWWordOT/CreatEXcelQWWordOT/bin/Debug/Form.xlsx"
            L = L.Replace(@"\", "/");
            string xlFileName = L;
            xlWB = App.Workbooks.Open(L);
            int StartIndex = 1;
          
            for (int i = 0; i <= Start.KolPov; i++)
            {

                try
                {
                int pp = StartIndex;
                Excel.Worksheet worksheet1 = App.Worksheets[s[i]];
                Excel.Worksheet worksheet2 = App.Worksheets["Отчет"];
                worksheet2.Columns.AutoFit();
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
                //Excel.Range Head27 = worksheet2.Range[worksheet2.Cells[2][StartIndex], worksheet2.Cells[6][StartIndex]];
                //Head27.Merge();
                StartIndex += 1;
                worksheet2.Cells[1][StartIndex] = worksheet1.Cells[1][7];
                    worksheet2.Cells[2][StartIndex] = worksheet1.Cells[2][7];
                StartIndex += 1;


                    //MessageBox.Show(StartIndex.ToString());
                    if (worksheet1.Cells[1][8].Formula == "Прочность резьбового соединения" )
                {
                        //MessageBox.Show(worksheet1.Cells[2][StartIndex - 7].Formula);

                        worksheet2.Cells[1][StartIndex] = worksheet1.Cells[1][8];
                    int Nwe = StartIndex;
                    StartIndex += 1;
                    worksheet2.Cells[2][StartIndex] = worksheet1.Cells[2][8];
                    StartIndex += 1;
                    int cla = StartIndex;
                    Excel.Range Head = worksheet2.Range[worksheet2.Cells[1][Nwe], worksheet2.Cells[1][cla]];
                    Head.Merge();
                    Excel.Range Head2 = worksheet2.Range[worksheet2.Cells[2][Nwe], worksheet2.Cells[2][cla]];
                    Head2.Merge();
                    worksheet2.Cells[3][Nwe] = "1)";
                    worksheet2.Cells[3][Nwe + 1] = "2)";
                    worksheet2.Cells[3][Nwe + 2] = "3)";
                    StartIndex += 1;
                    //MessageBox.Show(StartIndex.ToString()) ;
                }
                else
                {
                    worksheet2.Cells[1][StartIndex] = worksheet1.Cells[1][8];
                    worksheet2.Cells[2][StartIndex] = worksheet1.Cells[2][8];
                    StartIndex += 1;
                }

                    //MessageBox.Show(worksheet1.Cells[2][StartIndex - 15].Formula);
                    //MessageBox.Show(worksheet1.Cells[2][StartIndex - 13].Formula);
                    //MessageBox.Show(worksheet1.Cells[2][StartIndex - 14].Formula);
                    //MessageBox.Show(worksheet1.Cells[2][StartIndex - 12].Formula);
                    //MessageBox.Show(worksheet1.Cells[2][StartIndex - 11].Formula);
                    //MessageBox.Show(worksheet1.Cells[2][StartIndex - 10].Formula);
                    //MessageBox.Show(worksheet1.Cells[2][StartIndex - 9].Formula);
                    //MessageBox.Show(worksheet2.Cells[2][StartIndex-8]);
                    //MessageBox.Show(worksheet1.Cells[2][StartIndex - 7].Formula);

                    if (worksheet1.Cells[1][11].Formula == "Скручивание резьбы фитинга")
                {
                  
                    worksheet2.Cells[1][StartIndex] = worksheet1.Cells[1][11];
                    int Nwe = StartIndex;
                    StartIndex += 1;
                    worksheet2.Cells[2][StartIndex] = worksheet1.Cells[2][11];
                    StartIndex += 1;
                    int cla = StartIndex;
                    Excel.Range Head = worksheet2.Range[worksheet2.Cells[1][Nwe], worksheet2.Cells[1][cla]];
                    Head.Merge();
                    Excel.Range Head2 = worksheet2.Range[worksheet2.Cells[2][Nwe], worksheet2.Cells[2][cla]];
                    Head2.Merge();
                    worksheet2.Cells[3][Nwe] = "1)";
                    worksheet2.Cells[3][Nwe + 1] = "2)";
                    worksheet2.Cells[3][Nwe + 2] = "3)";
                        StartIndex += 1;


                    }
                else
                {

                    worksheet2.Cells[1][StartIndex] = worksheet1.Cells[1][9];
                    worksheet2.Cells[2][StartIndex] = worksheet1.Cells[2][9];
                    StartIndex += 1;
                }

                if (worksheet1.Cells[1][StartIndex].Formula == "Наружный диаметр бурта")
                {
                    worksheet2.Cells[1][StartIndex] = worksheet1.Cells[1][10];
                    worksheet2.Cells[2][StartIndex] = worksheet1.Cells[2][10];
                    StartIndex += 1;
                }
                else { }
                    //MessageBox.Show(StartIndex.ToString());
                    if (worksheet1.Cells[1][14].Formula == "Масса, г")
                    {
                        worksheet2.Cells[1][StartIndex] = worksheet1.Cells[1][14];
                        worksheet2.Cells[2][StartIndex] = worksheet1.Cells[2][14];
                        StartIndex += 1;
                        StartIndex += 1;
                    }
                    else { }
                    if (worksheet2.Cells[1][StartIndex].Formula == "Масса, г")
                {
                    worksheet2.Cells[1][StartIndex] = worksheet1.Cells[1][14];
                    worksheet2.Cells[2][StartIndex] = worksheet1.Cells[2][14];
                    StartIndex += 1;
                }
                    else { }
                   




                    int gg = StartIndex;

                    if (worksheet1.Cells[1][8].Formula == "Прочность резьбового соединения")
                    {
                        Excel.Range RR1 = worksheet2.Range[worksheet2.Cells[1][pp], worksheet2.Cells[6][gg-2]];
                        RR1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                            RR1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                            RR1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                            RR1.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                            RR1.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                            RR1.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                        StartIndex += 1;
                    }
                    else
                    {
                        Excel.Range RR1 = worksheet2.Range[worksheet2.Cells[1][pp], worksheet2.Cells[6][gg-4]];
                        RR1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                            RR1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                            RR1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                            RR1.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                            RR1.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                            RR1.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                        StartIndex += 1;
                    }
                    if (worksheet1.Cells[1][7].Formula == "Масса, г")
                    {
                        Excel.Range RR1 = worksheet2.Range[worksheet2.Cells[1][pp], worksheet2.Cells[6][gg-3]];
                        RR1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                            RR1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                            RR1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                            RR1.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                            RR1.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                            RR1.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                        StartIndex += 1;
                    }
                    else { }
                    if (worksheet1.Cells[1][7].Formula == "Внутренний Ø в месте присоединения к закладной (SIZE 5)")
                    {
                        Excel.Range RR1 = worksheet2.Range[worksheet2.Cells[1][pp], worksheet2.Cells[6][gg]];
                        RR1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                            RR1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                            RR1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                            RR1.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                            RR1.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                            RR1.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                        StartIndex += 1;
                    }
                    if (worksheet1.Cells[1][7].Formula == "Высота, мм")
                    {
                        Excel.Range RR1 = worksheet2.Range[worksheet2.Cells[1][pp], worksheet2.Cells[6][gg+2]];
                        RR1.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                            RR1.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                            RR1.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                            RR1.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                            RR1.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                            RR1.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
                        StartIndex += 1;
                    }
                }

                catch 
                {
                    MessageBox.Show(" Ошибка!!! Перезапустите приложение");
                }

               
            }
           
            App.Visible = true;
    


        }



        }
    }
