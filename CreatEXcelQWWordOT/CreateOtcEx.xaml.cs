
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
            if (BP114f_Copy3.IsChecked == true)
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
            if (F8.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист24";
            }
            if (F9.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист25";
            }
            if (F10.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист26";
            }
            if (F11.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист27";
            }
            if (F12.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист28";
            }
            if (F13.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист29";
            }
            if (F14.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист30";
            }
            if (F15.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист31";
            }
            if (F16.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист32";
            }
            if (F17.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист33";
            }
            if (F18.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист34";
            }
            if (F19.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист35";
            }
            if (F20.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист36";
            }
            if (F21.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист37";
            }
            if (F22.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист38";
            }
            if (F23.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист39";
            }
            if (F24.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист40";
            }
            if (F25.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист41";
            }
            if (F26.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист42";
            }
            if (F27.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист43";
            }
            if (F28.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист44";
            }
            if (F29.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист45";
            }
            if (F30.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист46";
            }
            if (F31.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист47";
            }
            if (F32.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист48";
            }

            MessageBox.Show(Start.KolPov.ToString());

            //MessageBox.Show(Start.str[1].ToString());
            Func.Viz(Start.str);
            


        }

  
    }
}
