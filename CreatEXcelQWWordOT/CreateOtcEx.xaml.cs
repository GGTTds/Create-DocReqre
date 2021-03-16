﻿
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using MessageBox = System.Windows.MessageBox;
using Path = System.IO.Path;

namespace CreatEXcelQWWordOT
{
    /// <summary>
    /// Логика взаимодействия для CreateOtcEx.xaml
    /// </summary>
    public partial class CreateOtcEx : Window
    {
        public CreateOtcEx()
        {
            Start.KolPov = -1;
            Start.StartIndex = 1;
            Start.LolPov = 1;
            InitializeComponent();
           

        }

        private void Get_Click(object sender, RoutedEventArgs e)
        {


            MainWindow WW = new MainWindow();
            WW.Show();
            this.Close();
            if (BP114f.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист2";

            }
            if (BP2f.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист3";


            }
            if (ZN3_4.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист4";


            }
            if(BP114f_Copy.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист5";
            }
            if (BP2f_Copy.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист6";
            }
            if (ZN3_4_Copy.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист7";
            }
            if(BP114f_Copy1.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист8";
            }
            if(BP2f_Copy1.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист9";
            }
            if (BP114f_Copy2.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист10";
            }
            if (BP114f_Copy2.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист11";
            }
            if (BP2f_Copy2.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист12";
            }
            if (ZN3_4_Copy2.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист13";
            }
            if (ZN3_4_Copy3.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист14";
            }
            if (BP2f_Copy3.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист15";
            }
            if (ZN3_4_Copy4.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист16";
            }
            if (F1.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист17";
            }
            if (F2.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист18";
            }
            if (F3.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист19";
            }
            if (F4.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист20";
            }
            if (F5.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист21";
            }
            if (F6.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист22";
            }
            if (F7.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист23";
            }
            MessageBox.Show(Start.KolPov.ToString());

            //MessageBox.Show(Start.str[1].ToString());
            Func.Viz(Start.str);
            


        }

  
    }
}
