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
using Newtonsoft.Json;
using UITest.Model;
using UITest.Util;

namespace UITest
{

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public static DataTable dt = new DataTable();
        public MainWindow()
        {
            
            InitializeComponent();

            Main.Content = new MainContent();
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
            Button2.IsEnabled = false;
            Task task = new Task(() =>
            {
                try
                {
                    this.Dispatcher.Invoke(DispatcherPriority.Normal, (ThreadStart)delegate ()
                    {
                        
                        Main.Content = new UserControl1();
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
                    Button2.IsEnabled = true;
                });
            });
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
                        Main.Content = new MainContent();
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
            
            Main.Content = new userControl.MyChart();
        }
    }
}
