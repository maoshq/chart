using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
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
using System.Windows.Threading;
using HandyControl.Controls;
using Newtonsoft.Json;
using UITest.Model;
using UITest.Util;
using MessageBox = System.Windows.MessageBox;
using Window = System.Windows.Window;

namespace UITest
{

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public static DataTable dt = new DataTable();
        public static userControl.MyChart myChart;
        public List<string> lists;
        public MainWindow()
        {
            
            InitializeComponent();
            myChart = new userControl.MyChart();

            Util.Tool.InitSetting();
            //Main.Children.Add(myChart);
        }


        private void StackPanel_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                this.DragMove();
            }
        }

        private void WrapPanel_MouseDown(object sender, MouseButtonEventArgs e)
        {
            /*            if (this.WindowState == WindowState.Maximized)
                        {
                            this.WindowState = WindowState.Normal; 
                        }
                        else
                        {
                            this.WindowState = WindowState.Maximized; 
                        }*/
        }
        private void btn_min_Click(object sender, RoutedEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            Main.Visibility = Visibility.Hidden;
            Main1.Visibility = Visibility.Visible;
            Main1.Content = new UserControl1();

        }
        

        private void ReturnMain(object sender, RoutedEventArgs e)
        {
            Button1.IsEnabled = false;
            Task task = new Task(() =>
            {
                try
                {
                    this.Dispatcher.Invoke(DispatcherPriority.Normal, (ThreadStart)delegate ()
                    {
                        Main1.Visibility = Visibility.Hidden;
                        Main.Visibility = Visibility.Visible;
                    });
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                    throw;
                }
            });
            task.Start();
            Task cwt = task.ContinueWith(t =>
            {
                this.Dispatcher.Invoke(DispatcherPriority.Normal, (ThreadStart)delegate ()
                {
                    Button1.IsEnabled = true;
                });
            });
        }

        private void Button2_Copy_Click(object sender, RoutedEventArgs e)
        {

            Main.Visibility = Visibility.Hidden;
            
            Main1.Content = new MainContent();
        }

        private void SideMenu_SelectionChanged(object sender, HandyControl.Data.FunctionEventArgs<object> e)
        {
            SideMenuItem sideMenuItem = e.Info as SideMenuItem;
            string header = sideMenuItem.Header.ToString();
            if (header == "Partner Analyze")
            {
                Main1.Visibility = Visibility.Hidden;
                Main.Visibility = Visibility.Visible;
                
            }
        }

        private void SideMenuItem_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            SideMenuItem sideMenuItem = e.Source as SideMenuItem;
            string header = sideMenuItem.Header.ToString();
            if (header == "Partner Analyze")
            {
                Main1.Visibility = Visibility.Hidden;
                Main.Visibility = Visibility.Visible;

            }
        }

        private void Back_Click(object sender, RoutedEventArgs e)
        {

            if (Main.CanGoBack)
            {
                Main.GoBack();
            }
            
        }


    }
}
