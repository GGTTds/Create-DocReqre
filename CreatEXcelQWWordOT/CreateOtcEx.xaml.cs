
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
            Start.KolPov = 1;
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
            //MessageBox.Show(Start.str[1].ToString());
            Func.Viz(Start.str);


        }

  
    }
}
