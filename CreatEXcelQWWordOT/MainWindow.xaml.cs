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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace CreatEXcelQWWordOT
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void CreaOtch_Click(object sender, RoutedEventArgs e)
        {
            CreateOtcEx WW = new CreateOtcEx();
            WW.Show();
            this.Close();
        }

        private void Edit_Click(object sender, RoutedEventArgs e)
        {
            EditProd WW = new EditProd();
            WW.Show();
            this.Close();
        }

        private void Add_Click(object sender, RoutedEventArgs e)
        {
            AddPro WW = new AddPro();
            WW.Show();
            this.Close();
        }
    }
}
