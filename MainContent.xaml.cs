using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices.ComTypes;
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

namespace UITest
{
    /// <summary>
    /// MainContent.xaml 的交互逻辑
    /// </summary>
    public partial class MainContent : UserControl
    {

        public bool OSflag = false;
        public bool VerFlag = false;
        

        public MainContent()
        {
            InitializeComponent();
            InitUI();
            Util.Tool.InitSetting();
        }

        public void InitUI()
        {
            combox1.Items.Add("Insider | Fe");
            combox1.Items.Add("Insider | Mn");
            combox1.Items.Add("2004 | Vb");
            combox1.Items.Add("1909 | 19H2");
            combox1.Items.Add("1903 | 19H1");
            combox1.Items.Add("1809 | RS5");
            combox1.Items.Add("1803 | RS4");
            combox1.Items.Add("1709 | RS3");
            combox1.Items.Add("1607 | RS1");

            //TextBox1.Text = "rltkapou64.dll";
            //,rltkapo64.dll,igdkmd64.sys
        }

        private void CheckComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            System.Collections.IList selectedItems = combox1.SelectedItems;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Task();
        }

        public void Task()
        {
            Loading1.Visibility = Visibility.Visible;
            Button1.IsEnabled = false;
            string[] Name;
            Dispatcher x = Dispatcher.CurrentDispatcher;

            if (TextBox1.Text.Contains(","))
            {
                Name = TextBox1.Text.Split(",");
            }
            else
            {
                Name = new string[] { TextBox1.Text };
            }
            Dictionary<string, string[]> condition = new Dictionary<string, string[]> { };
            Dictionary<string, string> arg = new Dictionary<string, string> { };

            if (combox1.SelectedItem != null)
            {
                string[] selectedItem = new string[combox1.SelectedItems.Count];
                for (int i = 0; i < combox1.SelectedItems.Count; i++)
                {
                    selectedItem[i] = combox1.SelectedItems[i].ToString();
                }
                condition.Add("ReleaseVersion", selectedItem);
            }
            else
            {
                //condition.Add("ReleaseVersion", new string[]{ "2004 | Vb"});
            }
            if (combox2.SelectedItem != null)
            {
                string[] selectedItem = new string[combox2.SelectedItems.Count];
                for (int i = 0; i < combox2.SelectedItems.Count; i++)
                {
                    selectedItem[i] = combox2.SelectedItems[i].ToString();
                }
                condition.Add("OSVersion", selectedItem);
            }
            if (combox3.SelectedItem != null)
            {
                string[] selectedItem = new string[combox3.SelectedItems.Count];
                for (int i = 0; i < combox3.SelectedItems.Count; i++)
                {
                    selectedItem[i] = combox3.SelectedItems[i].ToString();
                }
                condition.Add("DriverVersion", selectedItem);
            }

            bool? isChecked = new UserControl1().Model1.IsChecked;
            string arg1 = new UserControl1().Model1.IsChecked == true ? "NET" : "LOCAL";
            arg.Add("model", arg1);
            try
            {
                Task task = new Task(() =>
                {
                    try
                    {
                        Util.Tool.GenerateExcel(Name, arg, condition);
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
                        Loading1.Visibility = Visibility.Hidden;
                        Button1.IsEnabled = true;
                        if (check1.IsChecked.HasValue)
                        {
                            if (check1.IsChecked.Value == true)
                            {
                                if (MessageBox.Show("generate chart ?", "finish", MessageBoxButton.OKCancel) == MessageBoxResult.OK)
                                {
                                    Util.Tool.GenerateChart1(Name);
                                }
                            }
                            else
                            {
                                MessageBox.Show("Finish");
                            }
                        }

                    });
                });
            }
            catch (Exception)
            {
                throw;
}
}
        private void TextBox1_TextChanged(object sender, TextChangedEventArgs e)
        {
            VerFlag = true;
        }

        private void combox2_MouseEnter(object sender, MouseEventArgs e)
        {
            try
            {
                if (!OSflag)
                {
                    Task task = new Task(() =>
                    {
                        List<string> items = Util.Tool.GetItems("rltkapou64.dll", "OSVersion","NET");
                        this.Dispatcher.Invoke(DispatcherPriority.Normal, (ThreadStart)delegate ()
                        {
                            foreach (var item in items)
                            {
                                combox2.Items.Add(item);
                            }
                        });
                        OSflag = true;
                    });
                    task.Start();
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void combox3_MouseEnter(object sender, MouseEventArgs e)
        {
            try
            {
                List<string> lists = null;
                string text = TextBox1.Text;
                if (VerFlag)
                {
                    Task task = new Task(() =>
                    {
                        lists = Util.Tool.GetItems(text, "DriverVersion", "NET");                      
                    });
                    task.Start();
                    Task cwt = task.ContinueWith(t =>
                    {
                        this.Dispatcher.Invoke(DispatcherPriority.Normal, (ThreadStart)delegate ()
                        {
                            combox3.Items.Clear();
                            foreach (var item in lists)
                            {
                                combox3.Items.Add(item);
                            }
                        });
                        VerFlag = false;
                    });
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void initi(object sender, EventArgs e)
        {

        }

        private void GDR_Click(object sender, RoutedEventArgs e)
        {
            Hyperlink link = sender as Hyperlink;
            Process.Start(new ProcessStartInfo(link.NavigateUri.AbsoluteUri));
            
        }

        private void TextBox1_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            Pop.IsOpen = true;//设置为打开状态
        }
    }
}
