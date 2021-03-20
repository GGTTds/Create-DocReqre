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
using Path = System.IO.Path;
using  System.IO;
using Microsoft.Win32;
using System.Windows.Forms;
using OpenFileDialog = System.Windows.Forms.OpenFileDialog;
using MessageBox = System.Windows.Forms.MessageBox;
using Excel = Microsoft.Office.Interop.Excel;

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
            Fail();
            //System.Windows.MessageBox.Show(Start.PutinBB.ToString());
        }

        private void CreaOtch_Click(object sender, RoutedEventArgs e)
        {
            CreateOtcEx WW = new CreateOtcEx();
            WW.Show();
            this.Close();
        }

        private void Edit_Click(object sender, RoutedEventArgs e)
        {
            var App = new Excel.Application();
            Excel.Workbook xlWB;

            string L;
            StreamReader rr = new StreamReader("Put.txt");
            L = rr.ReadLine();
            
            L = L.Replace(@"\", "/");
            string xlFileName = L;
            xlWB = App.Workbooks.Open(L);
            App.Visible = true;
        }

        private void Add_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string str = dialog.FileName;
                Start.PutinBB = str;
                //System.Windows.MessageBox.Show(Start.PutinBB.ToString());
                StreamWriter ss = new StreamWriter("Put.txt");
                ss.WriteLine(Start.PutinBB.ToString());
                ss.Close();
                DialogResult dialogResult = MessageBox.Show("Путь к файлу задан", "Файл", MessageBoxButtons.OK);
                if (dialogResult == System.Windows.Forms.DialogResult.OK)  { Fail(); }



            }
        
        
        
        }


        public void Fail()
        {
            if (Start.PutinBB != null)
            { fr.Content = " Выбран "; }
            else
            { fr.Content = " Не выбран"; }
        }

       
    }
}
