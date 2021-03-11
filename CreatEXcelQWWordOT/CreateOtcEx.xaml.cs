using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace CreatEXcelQWWordOT
{
    /// <summary>
    /// Логика взаимодействия для CreateOtcEx.xaml
    /// </summary>
    public partial class CreateOtcEx : Window
    {
        public CreateOtcEx()
        {
            Start.Zav = " ";
            Start.KolTov = " ";
            Start.StartIndex = 1;
            Start.LolPov = 1;
            InitializeComponent();
   
        }

        private void Get_Click(object sender, RoutedEventArgs e)
        {
            if (BP114f.IsChecked == true)
            {
                Start.KolTov = KOlitictvo.Text.ToString();
                Start.Zav = Zavod.Text.ToString();
                Func.BP114F();
            }



        }
   



    }
}
