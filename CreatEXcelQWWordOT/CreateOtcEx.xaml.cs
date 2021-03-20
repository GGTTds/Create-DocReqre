
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


            ChekPress();


        }

        private void Window_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if(e.Key.Equals(Key.Enter) == true)
            {
                ChekPress();
            }
            if (e.Key.Equals(Key.Escape) == true)
            {
                MainWindow ww = new MainWindow();
                ww.Show();
                this.Close();
                
            }
        }
    
    public void ChekPress()
        {
          
            if (BP114f.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист2";

            }
            if (F87.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист104";
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
            if (BP114f_Copy.IsChecked == true)
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
            if (BP114f_Copy1.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист8";
            }
            if (BP2f_Copy1.IsChecked == true)
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
            if (F33.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист49";
            }
            if (F34.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист50";
            }
            if (F35.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист51";
            }
            if (F36.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист52";
            }
            if (F37.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист53";
            }
            if (F38.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист54";
            }
            if (F39.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист55";
            }
            if (F40.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист56";
            }
            if (F41.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист58";
            }
            if (F42.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист57";
            }
            if (F43.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист60";
            }
            if (F44.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист61";
            }
            if (F45.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист62";
            }
            if (F46.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист63";
            }
            if (F47.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист64";
            }
            if (F48.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист65";
            }
            if (F49.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист66";
            }
            if (F50.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист67";
            }
            if (F51.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист68";
            }
            if (F52.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист69";
            }
            if (F53.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист70";
            }
            if (F54.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист71";
            }
            if (F55.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист72";
            }
            if (F56.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист73";
            }
            if (F57.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист74";
            }
            if (F58.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист75";
            }
            if (F59.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист76";
            }
            if (F60.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист77";
            }
            if (F61.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист78";
            }
            if (F62.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист79";
            }
            if (F63.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист80";
            }

            if (F64.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист81";
            }
            if (F65.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист82";
            }
            if (F66.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист83";
            }
            if (F67.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист84";
            }
            if (F68.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист85";
            }
            if (F69.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист86";
            }
            if (F70.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист87";
            }
            if (F71.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист88";
            }
            if (F72.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист89";
            }
            if (F73.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист90";
            }
            if (F74.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист91";
            }
            if (F75.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист92";
            }
            if (F76.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист93";
            }
            if (F77.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист94";
            }
            if (F78.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист95";
            }
            if (F79.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист96";
            }
            if (F80.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист97";
            }
            if (F81.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист98";
            }
            if (F82.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист99";
            }
            if (F83.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист100";
            }
            if (F84.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист101";
            }
            if (F85.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист102";
            }
            if (F86.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист103";
            }

            if (F88.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист105";
            }
            if (F89.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист106";
            }
            if (F90.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист107";
            }
            if (F91.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист108";
            }
            if (F92.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист109";
            }
            if (F93.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист110";
            }
            if (F94.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист111";
            }
            if (F95.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист112";
            }
            if (F96.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист113";
            }
            if (F97.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист115";
            }
            if (F98.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист116";
            }
            if (F99.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист117";
            }
            if (F100.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист118";
            }
            if (F101.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист119";
            }
            if (F102.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист120";
            }
            if (F103.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист121";
            }
            if (F104.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист122";
            }
            if (F105.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист123";
            }
            if (F106.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист124";
            }
            if (F107.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист125";
            }
            if (F108.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист126";
            }
            if (F109.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист127";
            }
            if (F110.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист128";
            }
            if (F111.IsChecked == true)
            {
                Start.KolPov += 1;
                Start.str[Start.KolPov] = "Лист129";
            }
            MainWindow WW = new MainWindow();
            WW.Show();
            this.Close();
            MessageBox.Show(Start.KolPov.ToString());

            if (Start.KolPov == -1)
            {
                MessageBox.Show(" Вы нечего не выбрали!!!");

            }
            else
            {
                //MessageBox.Show(Start.str[1].ToString());
                Func.Viz(Start.str);
            }
        }
    
    
    
    }
}
