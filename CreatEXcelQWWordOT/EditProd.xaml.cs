﻿using System;
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

namespace CreatEXcelQWWordOT
{
    /// <summary>
    /// Логика взаимодействия для EditProd.xaml
    /// </summary>
    public partial class EditProd : Window
    {
        public EditProd()
        {
            InitializeComponent();
           
        }

        public void Button_Click(object sender, RoutedEventArgs e)
        {
            if (Bp114f.IsChecked == true)
            {
                string g = "ВР114F";
                ToNextEdit WW = new ToNextEdit(g);
                WW.Show();
                this.Close();
            }
        }
    }
}
