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
using System.Collections;

namespace CreatEXcelQWWordOT
{
    /// <summary>
    /// Логика взаимодействия для ToNextEdit.xaml
    /// </summary>
    public partial class ToNextEdit : Window
    {
  
        public  ToNextEdit(string r)
        {
            Start.Abje = r;
            InitializeComponent();
            if (Start.Abje == "ВР114F") 
            { 
                

                //pos.Text = DataElemProduc.ВР114F.PosOtv.ToString();
            

            
            }

            
            
        }
        
        private void Go_Click(object sender, RoutedEventArgs e)
        {
            if (Start.Abje == "ВР114F")
            { 
            // DataElemProduc.ВР114F.PosOtv = pos.Text.ToString(); MessageBox.Show(DataElemProduc.ВР114F.PosOtv);
            //    this.Close();
            }
            
        }
    
    
    public void TT()
        {
            if (Start.Abje == "ВР114F")
            {
                
            }
        }
    
    
    
    
    
    
    }
}
