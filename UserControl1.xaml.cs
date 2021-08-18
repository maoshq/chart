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
            List<string> lists = new List<string> { "reaktek32.dll","dtsdek.dll","list.dll", "example2.sys", "ifhlt.dll" };
            //new Model.Driver("reaktek32.dll");
            listBox.ItemsSource = lists;
            
            //listBox.Template.Triggers.Clear();   

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

        private void SearchBar_SearchStarted(object sender, HandyControl.Data.FunctionEventArgs<string> e)
        {

            if (e.Info == null)
            {
                return;
            }
            foreach (var driver in listBox.Items)
            {
                ListBoxItem listBoxItem = listBox.ItemContainerGenerator.ContainerFromItem(driver) as ListBoxItem;
                
                listBoxItem?.Show(driver.ToString().Contains(e.Info.ToLower()));
                if (listBox.SelectedItems.Contains(driver))
                {
                    listBoxItem?.Show(true);
                }
            }
            

        }

        private void checkListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
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
        private void chbxAll_Checked(object sender, RoutedEventArgs e)
        {
            listBox.SelectAll();
        }

        private void chbxAll_Unchecked(object sender, RoutedEventArgs e)
        {
            listBox.UnselectAll();
        }
        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            CheckBox checkBox = sender as CheckBox;
        }

    }
}
