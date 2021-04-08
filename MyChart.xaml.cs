using HandyControl.Tools.Extension;
using LiveCharts;
using LiveCharts.Configurations;
using LiveCharts.Wpf;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
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
using UITest.Model;
using UITest.Util;

namespace UITest.userControl
{
    /// <summary>
    /// MyChart.xaml 的交互逻辑
    /// </summary>
    public partial class MyChart : UserControl, INotifyPropertyChanged
    {
        public bool OSflag = false;
        public bool NameFlag = false;
        public bool BoxStatus = false;
        public bool Reflag = false;
        public bool Driverflag = false;

        public List<string> Osversion;

        public ObservableCollection<string> lists;
        public MyChart()
        {
            InitializeComponent();
            Util.Tool.InitSetting();
            Chart1.Navigate(new Uri(Directory.GetCurrentDirectory() + "/chart1.html"));
            this.Chart1.ObjectForScripting = new OprateBasic(this);

            //Uri uri = new Uri("chart1.html", UriKind.Relative);
            //Stream source = Application.GetResourceStream(uri).Stream;

            //Chart1.NavigateToStream(source);

            SeriesCollection = new SeriesCollection
            {
                new LineSeries
                {    
                    Values = new ChartValues<double> { }                  
                }
            };

            SeriesCollection1 = new SeriesCollection
              {
                  new ColumnSeries
                  {
                      Title = "1988",
                      Values = new ChartValues<double> { 10, 50, 39, 50, 5, 10 }
                  }
              };

            //adding series will update and animate the chart automatically
            SeriesCollection1.Add(new ColumnSeries
            {
                Title = "1989",
                Values = new ChartValues<double> { 12, 71, 41, 21, 9, 6 }
            });
            SeriesCollection2 = new SeriesCollection
            {
                new ColumnSeries
                {
                    DataLabels = true,
                    Values = new ChartValues<double> { }
                }
            };
            string currentDriver = JsonConvert.DeserializeObject<Settings>(File.ReadAllText(Environment.CurrentDirectory + "\\setting.json")).CurrentDriver;
            ColumnLabels = new[]
                 {
                     1, 2, 3, 4, 5, 6
                 };
            //ColumnLabels = new[] { "Maria", "Susan", "Charles" };

            splitButton1.Content = currentDriver;
            listBox3.SelectedItem = currentDriver;
            System.Collections.IList selectedItems = listBox.SelectedItems;
       
            DataContext = this;
        }

        public SeriesCollection SeriesCollection { get; set; }
        public SeriesCollection SeriesCollection1 { get; set; }
        public SeriesCollection SeriesCollection2 { get; set; }
        public string[] Labels { get; set; }
        public string[] Labels2 { get; set; }
        public int[] ColumnLabels { get; set; }
        public Func<double, string> YFormatter { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;
        List<string> lists1;
        int i = 0;
        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            lists = new ObservableCollection<string>();
            SeriesCollection.Clear();
            int i = 0;
            listBox3.ItemsSource = lists;

            lists1 = Util.Tool.QueryItem("drivername");

            Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadDelegate(Loadname), lists1[i]);
        }
        private void Loadname(string name)
        {
            i++;
            if (lists1.Count != 0 && lists1.Count > i)
            {
                lists.Add(name);
                Dispatcher.BeginInvoke(DispatcherPriority.Background, new LoadDelegate(Loadname), lists1[i]);
            }
        }

        private delegate void LoadDelegate(string name);

        protected virtual void OnPropertyChanged(string propertyName = null)
        {
            if (PropertyChanged != null)
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private void searchBar_SearchStarted(object sender, HandyControl.Data.FunctionEventArgs<string> e)
        {
            if (e.Info == null)
            {
                chbxAll.Visibility = Visibility.Visible;
                return;
            }
            chbxAll.Visibility = Visibility.Collapsed;
            foreach (var driver in Osversion)
            {
                ListBoxItem listBoxItem = listBox.ItemContainerGenerator.ContainerFromItem(driver) as ListBoxItem;

                listBoxItem?.Show(driver.ToString().Contains(e.Info.ToLower()));
                if (listBox.SelectedItems.Contains(driver))
                {
                    listBoxItem?.Show(true);
                }
            }

        }

        private void chbxAll_Checked(object sender, RoutedEventArgs e)
        {
            listBox.SelectAll();
            
        }

        private void chbxAll_Unchecked(object sender, RoutedEventArgs e)
        {
            listBox.UnselectAll();
        }

        private void listBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (listBox.SelectedItems.Count == 1)
            {
                splitButton.Content = listBox.SelectedItem;
            }
            else if (listBox.SelectedItems.Count > 1)
            {
                splitButton.Content = "多选";
            }
            else if (listBox.SelectedItems.Count == 0)
            {
                splitButton.Content = "所有";
            }
            chbxAll.IsChecked = listBox.SelectedItems.Count == 0 ? false :
                    listBox.SelectedItems.Count == listBox.Items.Count ? (bool?)true : null;
        }
        private void listBox2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (listBox2.SelectedItems.Count == 1)
            {
                splitButton2.Content = listBox2.SelectedItem;
            }
            else if (listBox2.SelectedItems.Count > 1)
            {
                splitButton2.Content = "多选";
            }
            else if (listBox2.SelectedItems.Count == 0)
            {
                splitButton2.Content = "所有";
            }

        }
        private void splitButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                splitButton.IsDropDownOpen = true;
                chbxAll.Visibility = Visibility.Visible;
                if (!Reflag)
                {
                    Task task = new Task(() =>
                    {
                        Osversion = Util.Tool.QueryItem("osversion");
                        Osversion.Sort();
                        Osversion.RemoveRange(0, 308);
                    });
                    task.Start();
                    Task cwt = task.ContinueWith(t =>
                    {
                        this.Dispatcher.Invoke(DispatcherPriority.Background, (ThreadStart)delegate ()
                        {

                            foreach (var item in Osversion)
                            {
                                listBox.Items.Add(item);
                            }
                           
                           
                        });
                        Reflag = true;
                    });
                }   
            }
            catch (Exception)
            {
                throw;
            }
        }
        Dictionary<string, List<string>> driverData = new Dictionary<string, List<string>>();
        Dictionary<string, List<string>> data = new Dictionary<string, List<string>>();
        List<string> crash = new List<string>();
        string[] dateTimes;
        private void Button2_Click(object sender, RoutedEventArgs e)
        {
            Loading1.Visibility = Visibility.Visible;
            Button2.IsEnabled = false;
            ChartValues<int> crashes = new LiveCharts.ChartValues<int>();
            driverData.Clear();data.Clear();
            string currentDriver = JsonConvert.DeserializeObject<Settings>(File.ReadAllText(Environment.CurrentDirectory + "\\setting.json")).CurrentDriver;
            string osversion = "";
            string releaseversion = "";
            string driverversion = "";
            if (listBox2.SelectedItems.Count != 0)
            {
                List<string> list = new List<string>();
                foreach (var item in listBox2.SelectedItems)
                {
                    list.Add(item.ToString());
                }
                driverData.Add("ReleaseVersion", list);
            }
            if (listBox.SelectedItems.Count != 0)
            {
                List<string> list = new List<string>();
                foreach (var item in listBox.SelectedItems)
                {
                    list.Add(item.ToString());
                }
                driverData.Add("OSVersion", list);
            }
            if (combox3.SelectedItems.Count != 0)
            {
                List<string> list = new List<string>();
                foreach (var item in combox3.SelectedItems)
                {
                    list.Add(item.ToString());
                }
                driverData.Add("DriverVersion", list);
            }

            if (driverData.ContainsKey("ReleaseVersion") && driverData["ReleaseVersion"].Count != 0)
            {
                osversion = driverData["ReleaseVersion"].Count > 1 ? driverData.GetValueOrDefault("ReleaseVersion")[0] + "... " : driverData.GetValueOrDefault("ReleaseVersion")[0] + " ";
            }
            if (driverData.ContainsKey("OSVersion") && driverData["OSVersion"].Count != 0) 
            {
                releaseversion = driverData["OSVersion"].Count > 1 ? driverData.GetValueOrDefault("OSVersion")[0] + "... " : driverData.GetValueOrDefault("OSVersion")[0] + " ";
            }
            if (driverData.ContainsKey("DriverVersion") && driverData["DriverVersion"].Count != 0) 
            {
                driverversion = driverData["DriverVersion"].Count > 1 ? driverData.GetValueOrDefault("DriverVersion")[0] + "... " : driverData.GetValueOrDefault("DriverVersion")[0] + " ";
            }
            Task task = new Task(() =>
            {
                data = Util.Tool.QuerySet(currentDriver, driverData);

            });
            task.Start();
            
            Task cwt = task.ContinueWith(t =>
            {

                List<string> date = data.GetValueOrDefault("date");
                crash = data.GetValueOrDefault("crash");
                
                dateTimes = date.ToArray();
                foreach (var item in crash)
                {
                    crashes.Add(Convert.ToInt32(item));
                }
                this.Dispatcher.Invoke(DispatcherPriority.Normal, (ThreadStart)delegate ()
                {
                    SeriesCollection.Add(new LineSeries
                    {
                        Title = currentDriver + " " + osversion + releaseversion + driverversion,
                        Values = crashes,

                    });
                    AxesCollection axes = new AxesCollection()
                    {
                        new Axis()
                        {
                            Labels = dateTimes
                        }
                    };
                    Labels = dateTimes;

                    myChart.AxisX = axes;
                    DataContext = this;
                    Loading1.Visibility = Visibility.Hidden;
                    Button2.IsEnabled = true;
                });

            });
        }

        private void Button3_Click(object sender, RoutedEventArgs e)
        {
            SeriesCollection.Clear();
        }


        private void OS_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                splitButton2.IsDropDownOpen = true;
                if (!OSflag)
                {
                    Task task = new Task(() =>
                    {
                        Osversion = Util.Tool.QueryItem("releaseversion");
                        Osversion.Sort();
                        Osversion.RemoveRange(0, 2);
                        this.Dispatcher.Invoke(DispatcherPriority.Normal, (ThreadStart)delegate ()
                        {
                            foreach (var item in Osversion)
                            {
                                listBox2.Items.Add(item);
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
            List<string> lists = null;
            string text = listBox3.SelectedItem!=null ? listBox3.SelectedItem.ToString() : "";
            try
            {
                if (Driverflag)
                {
                    Task task = new Task(() =>
                    {
                        lists = Util.Tool.QueryItem("driverversion",text);
                        lists.Sort();
                        this.Dispatcher.Invoke(DispatcherPriority.Normal, (ThreadStart)delegate ()
                        {
                            combox3.Items.Clear();
                            foreach (var item in lists)
                            {
                                combox3.Items.Add(item);
                            }
                        });
                        Driverflag = false;
                    });
                    task.Start();
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        private async Task AsyncAccess()
        {
            var getDataListTask = new Task(() =>
            {
                //耗时的计算或请求等操作的代码写在这里
                Thread.Sleep(5000);
            });
            getDataListTask.Start();
            await getDataListTask;
            var fillModelTask = Task.Factory.StartNew(() =>
            {
                
            }, CancellationToken.None, TaskCreationOptions.None, TaskScheduler.FromCurrentSynchronizationContext());
            await fillModelTask;
        }
        private void splitButton1_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                splitButton1.IsDropDownOpen = true;
                if (!NameFlag)
                {
                    Task task = new Task(() =>
                    {
                        //lists = Util.Tool.QueryItem("drivername");

                        this.Dispatcher.Invoke(DispatcherPriority.Normal,new Action(()=>
                        {
                            

                        }));
                        NameFlag = true;
                    });
                    task.Start();
                }
            }
            catch (Exception)
            {
                throw;
            }
        }
        private void GDR_Click(object sender, RoutedEventArgs e)
        {
            Hyperlink link = sender as Hyperlink;

            System.Diagnostics.Process.Start("explorer", link.NavigateUri.ToString());
        }
        private void listBox3_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Driverflag = true;
            Settings model = JsonConvert.DeserializeObject<Settings>(File.ReadAllText(Environment.CurrentDirectory + "\\setting.json"));
            {
                if (listBox3.SelectedItem != null)
                {
                    model.CurrentDriver = listBox3.SelectedItem.ToString();
                }
            };
            File.WriteAllText(Tool.SettingPath, JsonConvert.SerializeObject(model, Formatting.Indented));

            if (listBox3.SelectedItems.Count == 1)
            {
                splitButton1.Content = listBox3.SelectedItem;
            }
            else if (listBox3.SelectedItems.Count == 0)
            {
                splitButton1.Content = "";
            }
        }

        private void searchBar1_SearchStarted(object sender, HandyControl.Data.FunctionEventArgs<string> e)
        {
            if (e.Info == null)
            {
                return;
            }
            foreach (var driver in lists)
            {
                ListBoxItem listBoxItem = listBox3.ItemContainerGenerator.ContainerFromItem(driver) as ListBoxItem;

                listBoxItem?.Show(driver.ToString().Contains(e.Info.ToLower()));
                if (listBox.SelectedItems.Contains(driver))
                {
                    listBoxItem?.Show(true);
                }
            }
        }

        private void export_Click(object sender, RoutedEventArgs e)
        {

        }

        private void sp_Clear_Click(object sender, RoutedEventArgs e)
        {
            if (listBox.HasItems)
            {
                listBox.UnselectAll();
            }
        }

        private void sp2_Clear_Click(object sender, RoutedEventArgs e)
        {
            if (listBox2.HasItems)
            {
                listBox2.UnselectAll();
            }
        }

        class DataGridSource
        {
            public String OS { set; get; }

            public String GDR { set; get; }

            public String DriverVersion { set; get; }

            public int crash { set; get; }

            public int ImpactedMachines { set; get; }

            public int TotalMachines { set; get; }

            public Decimal PercentImpacted { set; get; }
        }
        List<DataGridSource> sources ;
        string CurrentDriver;
        List<Object> list3; 
        private void click(object sender, ChartPoint chartPoint)
        {
            List<string> lable = new List<string>();
            sources = new List<DataGridSource>();
            DataGridSource ds = new DataGridSource();
            
            ChartValues<int> crashes = new LiveCharts.ChartValues<int>();
            object v = myChart.GetValue(DataContextProperty);
            Tab1.SelectedIndex = 3;
            SeriesCollection series = myChart.Series;
            int currentSeriesIndex = series.CurrentSeriesIndex;
            string CurrentDate = series.Chart.AxisX[0].Labels[Convert.ToInt32(chartPoint.X)];
            CurrentDriver = series[currentSeriesIndex - 1].Title.Split(" ")[0];

            DataTable dt = Tool.QueryTest(CurrentDriver, CurrentDate);
            
            List<object> lists = dt.AsEnumerable().Select(d => d[0]).ToList();
            List<object> lists1 = dt.AsEnumerable().Select(d => d[2]).ToList();
            lists.Sort();
            lists.Reverse();

            list3 = lists;

            DataView defaultView = dt.DefaultView;
            defaultView.Sort = "Crashes desc";
            DataTable dataTable = defaultView.ToTable();
            
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                DataGridSource wd = new DataGridSource
                {
                    OS = dataTable.Rows[i].ItemArray[2].ToString(),
                    GDR = dataTable.Rows[i].ItemArray[1].ToString(),
                    DriverVersion = dataTable.Rows[i].ItemArray[7].ToString(),
                    crash = Convert.ToInt32(dataTable.Rows[i].ItemArray[0]),
                    ImpactedMachines = Convert.ToInt32(dataTable.Rows[i].ItemArray[3]),
                    TotalMachines = Convert.ToInt32(dataTable.Rows[i].ItemArray[4]),
                    PercentImpacted = Convert.ToDecimal(dataTable.Rows[i].ItemArray[6])
            };
                sources.Add(wd);
            }

            DataGrid1.ItemsSource = sources;
            //DataGrid1.SetValue(DataGrid.StyleProperty, Application.Current.Resources["GridStyle"]);

            //EnumerableRowCollection<DataRow> enumerableRowCollections = dataTable.AsEnumerable().Where((r => r["Osversion"].ToString() == "2004 | Vb"));
            foreach (var item in lists)
            {
                crashes.Add(Convert.ToInt32(item));
            }
            if (crashes.Count > 50)
            {
                for (int i = crashes.Count-1; i > 50; i--)
                {
                    crashes.Remove(crashes[i]);
                    if (crash.Count==50)
                    {
                        crashes[50] += crashes[i];
                        crashes.Remove(crashes[i]);
                    }
                    
                }
            }

            
            foreach (var item in lists1)    
            {
                lable.Add(Convert.ToString(item));
            }
            SeriesCollection2.Clear();
            SeriesCollection2.Add(new ColumnSeries
            {
                Title = CurrentDriver,
              
                Values = crashes,
            });
            AxesCollection axes = new AxesCollection()
                    {
                        new Axis()
                        {
                            Labels = lable
                        }
                    };
            //Labels = lists1;
            myChart2.AxisX = axes;
            DataContext = this;
        }

        private void RadioButton_Click(object sender, RoutedEventArgs e)
        {
            List<string> lable = new List<string>();
            ChartValues<int> crashes = new LiveCharts.ChartValues<int>();
            if (sources.Count >0)
            {
                var list1 = (from l in sources
                           group l by new { l.GDR } into g
                           select new
                           {
                               g.Key.GDR,
                               crash = g.Sum(c => c.crash),
                               impactedMachine = g.Sum(c => c.ImpactedMachines),
                               TotalMachines = g.Sum(c => c.TotalMachines),
                               Percent = Decimal.Parse((Convert.ToDecimal(g.Sum(c => c.ImpactedMachines)) / Convert.ToDecimal(g.Sum(c => c.TotalMachines)) * 100).ToString("0.00"))
                           }).ToList().ToList().OrderByDescending(x => x.crash);

                foreach (var item in list1)
                {
                    crashes.Add(item.crash);
                }
                if (crashes.Count > 20)
                {
                    for (int i = crashes.Count - 1; i > 20; i--)
                    {
                        crashes[20] += crashes[i];
                        crashes.Remove(crashes[i]);
                    }
                }

                foreach (var item in list1)
                {
                    lable.Add(Convert.ToString(item.GDR));
                }
                SeriesCollection2.Clear();
                SeriesCollection2.Add(new ColumnSeries
                {
                    Values = crashes,
                    Title = CurrentDriver,
                });
                AxesCollection axes = new AxesCollection()
                    {
                        new Axis()
                        {
                            Labels = lable
                        }
                    };
                myChart2.AxisX = axes;
                DataContext = this;

                DataGrid1.ItemsSource = list1;
            }
            
        }

        private void RadioButton_Click_1(object sender, RoutedEventArgs e)
        {
            List<string> lable = new List<string>();
            ChartValues<int> crashes = new LiveCharts.ChartValues<int>();
            if (sources.Count > 0)
            {
                var list1 = (from l in sources
                             group l by new { l.OS } into g
                             select new
                             {
                                 g.Key.OS,
                                 crash = g.Sum(c => c.crash),
                                 impactedMachine = g.Sum(c => c.ImpactedMachines),
                                 TotalMachines = g.Sum(c => c.TotalMachines),
                                 Percent = Decimal.Parse((Convert.ToDecimal(g.Sum(c => c.ImpactedMachines)) / Convert.ToDecimal(g.Sum(c => c.TotalMachines)) * 100).ToString("0.00"))
                             }).ToList().ToList().OrderByDescending(x => x.crash);

                foreach (var item in list1)
                {
                    crashes.Add(item.crash);
                }
                if (crashes.Count > 20)
                {
                    for (int i = crashes.Count - 1; i > 20; i--)
                    {
                        crashes[20] += crashes[i];
                        crashes.Remove(crashes[i]);
                    }
                }

                foreach (var item in list1)
                {
                    lable.Add(Convert.ToString(item.OS));
                }
                SeriesCollection2.Clear();
                SeriesCollection2.Add(new ColumnSeries
                {
                    Values = crashes,
                    Title = CurrentDriver,
                });
                
                AxesCollection axes = new AxesCollection()
                    {
                        new Axis()
                        {
                            Labels = lable,
                            DisableAnimations = true
                        }
                    };
                myChart2.AxisX = axes;
                DataContext = this;

                DataGrid1.ItemsSource = list1;
            }
        }

        private void RadioButton_Click_2(object sender, RoutedEventArgs e)
        {
            List<string> lable = new List<string>();
            ChartValues<int> crashes = new LiveCharts.ChartValues<int>();
            if (sources.Count > 0)
            {
                var list1 = (from l in sources
                             group l by new { l.DriverVersion } into g 
                             select new
                             {
                                 g.Key.DriverVersion,
                                 crash = g.Sum(c => c.crash),
                                 impactedMachine = g.Sum(c => c.ImpactedMachines),
                                 TotalMachines = g.Sum(c => c.TotalMachines),
                                 Percent = Decimal.Parse((Convert.ToDecimal(g.Sum(c => c.ImpactedMachines)) / Convert.ToDecimal(g.Sum(c => c.TotalMachines)) * 100).ToString("0.00"))
                             }).ToList().OrderByDescending(x=>x.crash);
               

                foreach (var item in list1)
                {
                    crashes.Add(item.crash);
                }
              
                if (crashes.Count > 20)
                {
                    for (int i = crashes.Count - 1; i > 20; i--)
                    {
                        crashes[20] += crashes[i];
                        crashes.Remove(crashes[i]);
                    }
                }

                foreach (var item in list1)
                {
                    lable.Add(Convert.ToString(item.DriverVersion));
                }
                SeriesCollection2.Clear();
                SeriesCollection2.Add(new ColumnSeries
                {
                    Values = crashes,
                    Title = CurrentDriver,
                });

                AxesCollection axes = new AxesCollection()
                    {
                        new Axis()
                        {
                            Labels = lable
                        }
                    };
                myChart2.AxisX = axes;
                DataContext = this;

                DataGrid1.ItemsSource = list1;
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            string Crashes = JsonConvert.SerializeObject(list3);
            string XDate = JsonConvert.SerializeObject(dateTimes);
            Chart1.InvokeScript("GetData", Crashes, XDate);

        }
        public void Wtest(string str)
        {
            MessageBox.Show(str);
        }
    }
    [System.Runtime.InteropServices.ComVisible(true)]
    public class OprateBasic
    {
        private MyChart instance;
        public OprateBasic(MyChart instance)
        {
            this.instance = instance;
        }

        public void HandleTest(string p)
        {
            instance.Wtest(p);
        }
    }
}

