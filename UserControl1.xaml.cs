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
            if (File.Exists(Tool.SettingPath))
            {
                File.Delete(Tool.SettingPath);
            }
            var setting = new Settings()
            {
                isNet = ! Model1.IsChecked.HasValue,
            };
            File.WriteAllText(Tool.SettingPath, JsonConvert.SerializeObject(setting, Formatting.Indented));
        }

        private void ModelCheck2(object sender, RoutedEventArgs e)
        {
            if (File.Exists(Tool.SettingPath))
            {
                File.Delete(Tool.SettingPath);
            }
            var setting = new Settings()
            {
                isNet = Model1.IsChecked.HasValue,
            };
            File.WriteAllText(Tool.SettingPath, JsonConvert.SerializeObject(setting,Formatting.Indented));
        }
    }
}
