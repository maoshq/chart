using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace UITest.userControl
{
    /// <summary>
    /// Data1.xaml 的交互逻辑
    /// </summary>
    public partial class Data1 : UserControl
    {
        public Data1()
        {
            InitializeComponent();
     
            Web.Navigate(new Uri(Directory.GetCurrentDirectory() + "/Chart1.html"));
        }


    }
}
