using HandyControl.Tools.Extension;
using Newtonsoft.Json;
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
using UITest.Model;
using UITest.Util;
namespace UITest
{
    /// <summary>
    /// UserControl1.xaml 的交互逻辑
    /// </summary>
    public partial class UserControl1 : UserControl
    {

        public UserControl1()
        {
            InitializeComponent();

            this.SettingBinding();

            
            //listBox.Template.Triggers.Clear();   

            
        }
        public static T FindVisualChild<T>(DependencyObject obj) where T : DependencyObject
        {
            if (obj != null)
            {
                for (int i = 0; i < VisualTreeHelper.GetChildrenCount(obj); i++)
                {
                    DependencyObject child = VisualTreeHelper.GetChild(obj, i);
                    if (child != null && child is T)
                    {
                        return (T)child;
                    }
                    T childItem = FindVisualChild<T>(child);
                    if (childItem != null) return childItem;
                }
            }
            return null;
        }
        private void SettingBinding()
        {
            Tool.InitSetting();
            Settings model = JsonConvert.DeserializeObject<Settings>(File.ReadAllText(Environment.CurrentDirectory + "\\setting.json"));

            Binding binding = new Binding("isNet")
            {
                Source = model
            };
            Model1.SetBinding(RadioButton.IsCheckedProperty, binding);
            Model2.IsChecked = ! Model1.IsChecked.Value;


        }

        private void ModelCheck1(object sender, RoutedEventArgs e)
        {
            Settings model = JsonConvert.DeserializeObject<Settings>(File.ReadAllText(Environment.CurrentDirectory + "\\setting.json"));
            {
                model.isNet = !Model1.IsChecked.HasValue;
            };
            File.WriteAllText(Tool.SettingPath, JsonConvert.SerializeObject(model, Formatting.Indented));
        }

        private void ModelCheck2(object sender, RoutedEventArgs e)
        {
            Settings model = JsonConvert.DeserializeObject<Settings>(File.ReadAllText(Environment.CurrentDirectory + "\\setting.json"));
            {
                model.isNet = Model1.IsChecked.HasValue;
            };
            File.WriteAllText(Tool.SettingPath, JsonConvert.SerializeObject(model, Formatting.Indented));
        }

        private void Main_Loaded(object sender, RoutedEventArgs e)
        {


        }


    }
}
