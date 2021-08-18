using HandyControl.Tools.Extension;
using LiveCharts;
using LiveCharts.Configurations;
using LiveCharts.Wpf;
using System;
using System.Collections.Generic;
using System.ComponentModel;
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

        private bool _mariaSeriesVisibility;

        public static Dictionary<string, List<string>> driverData = new Dictionary<string, List<string>>();
        ChartValues<int> crashes = new LiveCharts.ChartValues<int>(); ChartValues<int> total = new ChartValues<int>(); ChartValues<int> tmad = new ChartValues<int>();

        public List<string> lists;
        public List<string> Osversion;
        public MyChart()
        {
            InitializeComponent();

            System.Collections.IList selectedItems = listBox.SelectedItems;
            /*CrashesVisibility = true;
            Dictionary<string, List<string>> data = Util.Tool.QuerySet("rltkapou64.dll","");
            <string> date = data.GetValueOrDefault("date");
            List<string> crash = data.GetValueOrDefault("crash");
            if (driverData.Count == 0)
            {
                *//*                driverData.Add("2020-11-20", new List<string> { "30024", "1254700","3960034" });
                                driverData.Add("2020-11-23", new List<string> { "36682", "1368270","3980034" });
                                driverData.Add("2020-11-25", new List<string> { "39573", "1375160", "3990034" });
                                driverData.Add("2020-11-30", new List<string> { "30024", "1322710", "4037633" });
                                driverData.Add("2020-12-3", new List<string> { "40294", "1292510", "4057633" });
                                driverData.Add("2020-12-6", new List<string> { "43809", "1310800", "4037633" });
                                driverData.Add("2020-12-11", new List<string> { "45336", "1293440", "4057333" });
                                driverData.Add("2020-12-16", new List<string> { "47277", "1303080", "4087633" });*//*
            }
            string[] dateTimes = date.ToArray();
            foreach (var item in crash)
            {
                crashes.Add(Convert.ToInt32(item));
            }

            SeriesCollection = new SeriesCollection
            {
                new LineSeries
                {
                    Title = "rltkapou64.dll",
                    Values = crashes,
                    //DataLabels = true,
                    PointGeometrySize = 5,
                    Fill = Brushes.Transparent,
                    Stroke = Brushes.Orange,
                    Visibility = 0
                },
*//*                new LineSeries
                {
                    Title = "TMAD",
                    Values = tmad,
                    //DataLabels = true,
                    PointGeometrySize = 5,
                    Fill = Brushes.Transparent,
                    Stroke = Brushes.Black,
                    ScalesYAt = 1
                },
                new LineSeries
                {
                    Title = "Total",
                    Values = total,
                    //DataLabels = true,
                    PointGeometrySize = 5,
                    Fill = Brushes.Transparent,
                    Stroke = Brushes.Red,
                    ScalesYAt = 1,
                    //Visibility = CrashesVisibility==true?0:Visibility.Hidden,
                },*//*
            };

            Labels = dateTimes;
*/

                    //YFormatter = value => value.ToString("C");

                    //modifying the series collection will animate and update the chart
                    /*            SeriesCollection.Add(new LineSeries
                                {
                                    Values = new ChartValues<double> { 5, 3, 2, 4 },
                                    LineSmoothness = 0 //straight lines, 1 really smooth lines
                                });*/

                    //modifying any series values will also animate and update the chart
                    //SeriesCollection[1].Values.Add(5d);

                    DataContext = this;


        }


        public SeriesCollection SeriesCollection { get; set; }
        public string[] Labels { get; set; }
        public Func<double, string> YFormatter { get; set; }
        public bool CrashesVisibility
        {
            get { return _mariaSeriesVisibility; }
            set
            {
                _mariaSeriesVisibility = value;
                OnPropertyChanged("CrashesVisibility");
            }
        }



        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName = null)
        {
            if (PropertyChanged != null)
                PropertyChanged.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private void searchBar_SearchStarted(object sender, HandyControl.Data.FunctionEventArgs<string> e)
        {
            if (e.Info == null)
            {
                return;
            }
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

        private void splitButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                splitButton.IsDropDownOpen = true;
                if (!OSflag)
                {
                    Task task = new Task(() =>
                    {
                        Osversion = Util.Tool.QueryItem("osversion");
                        Osversion.Sort();
                        this.Dispatcher.Invoke(DispatcherPriority.Normal, (ThreadStart)delegate ()
                        {
                            foreach (var item in Osversion)
                            {
                                listBox.Items.Add(item);
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

        private void combobox1_MouseEnter(object sender, MouseEventArgs e)
        {
            try
            {
                if (!NameFlag)
                {
                    Task task = new Task(() =>
                    {
                        lists = Util.Tool.QueryItem("drivername");
                        
                        this.Dispatcher.Invoke(DispatcherPriority.Normal, (ThreadStart)delegate ()
                        {
                            combobox1.ItemsSource = lists;
                        });
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
    }

}

