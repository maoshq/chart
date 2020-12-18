using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Linq.Dynamic;
using System.Linq.Dynamic.Core;
using System.Windows;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using System.Net;
using Newtonsoft.Json.Schema;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using UITest.Model;
namespace UITest.Util
{
    public class Tool
    {
        public const string NetPath = "\\\\172.30.184.28\\psd\\Common\\Auto Testing\\Auto Tools\\CrashesTool_v1.2\\ExportData\\";
        const string Crashes_Total = "Reliability-Crashes_Total-";
        const string TMAD = "OSAdoption-TMAD-";
        const string Crashes = "Reliability-Crashes-";
        const string ReportDate = "Reliability-Crashes_Date-";
        public static string Path = System.Environment.CurrentDirectory + "\\ExportData\\";
        public const string suffix = ".csv";
        public static readonly string[] Name = { "rltkapou64.dll", "rltkapo64.dll", "igdkmd64.sys" };//rltkapou64.dll,rltkapo64.dll,igdkmd64.sys
        public static readonly string[] OS = { "10.0.19041.508", "10.0.19041.572" };
        public static readonly string[] Release = { "2004 | Vb" };
        public static readonly string[] DriverVersion = { "26.20.100.7872" };
        public static readonly Dictionary<string, string[]> Condition = new Dictionary<string, string[]> {
            {"OSVersion",OS},
            {"ReleaseVersion",Release},
            {"DriverVersion",DriverVersion}
        };

        public static readonly string SettingPath = System.Environment.CurrentDirectory + "\\setting.json";
        public static void InitSetting()
        {
            /*            string schemaJson = @"{
                            'isNet': true
                        }";*/
            if (!File.Exists(System.Environment.CurrentDirectory + "\\setting.json"))
            {
                Settings settings = new Settings(true,new List<string> { "rltkapou64.dll" },false);
                string json = JsonConvert.SerializeObject(settings, Formatting.Indented);

                File.AppendAllText(System.Environment.CurrentDirectory + "\\setting.json", json, Encoding.UTF8);
            }
        }
        public static void QueryByJSON(string json,string key)
        {
            using StreamReader file = File.OpenText(System.Environment.CurrentDirectory + "\\setting.json");
            using JsonTextReader reader = new JsonTextReader(file);

            string Model = JsonConvert.DeserializeObject<Dictionary<string, string>>(json)[key];
            
        }

        public static void GenerateExcel(string[] name, Dictionary<string, string> arg, Dictionary<string, string[]> condition)
        {
            Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
            if (File.Exists(System.Environment.CurrentDirectory + "\\RecentData.xlsx"))
            {
                workbook.LoadFromFile(System.Environment.CurrentDirectory + "\\RecentData.xlsx");
            }
            else
            {
                InitiExcel(name);
                workbook.LoadFromFile(System.Environment.CurrentDirectory + "\\RecentData.xlsx");
            }
            LoadData(workbook, name, arg, condition);
            workbook.SaveToFile(System.Environment.CurrentDirectory + "\\RecentData.xlsx", ExcelVersion.Version2013);

            GenerateChart(name);
        }
        public static void LoadData(Spire.Xls.Workbook workbook, string[] name, Dictionary<string, string> arg, Dictionary<string, string[]> condition)
        {
            //bool flag = readCSV("\\\\172.30.184.28\\psd\\Common\\Auto Testing\\Auto Tools\\CrashesTool_v1.2\\ExportData\\Reliability-Crashes_Total-11-27.csv", out DataTable dt, "NET");
            Spire.Xls.Worksheet sheet = workbook.Worksheets["Crashes"];
            int lastRow = sheet.Range.LastRow;              //2020/10/12 0:00:00

            arg.TryGetValue("model", out string model);

            String path = model == "NET" ? NetPath : Path;
            //string[] fileName = Directory.GetFiles(System.Environment.CurrentDirectory + "\\ExportData", "*.csv");
            string[] fileName = Directory.GetFiles(path, "*.csv");
            string firstDate = DateTime.Now.Year.ToString() + "/" + fileName[0].Split("TMAD-")[1].Split(".")[0].Split("-")[1] + "/" + fileName[0].Split("TMAD-")[1].Split(".")[0].Split("-")[2]; //获取文件列表中第一项文件日期
            //string lastDate = DateTime.Now.Year.ToString() + "/" + fileName[^1].Split("Total-")[1].Split(".")[0].Split("-")[0] + "/" + fileName[^1].Split("Total-")[1].Split(".")[0].Split("-")[1];

            string[] date_1 = sheet.Range["B" + lastRow].Value == "Date" ? firstDate.Split("/") : sheet.Range["B" + lastRow].Value.Split(" ")[0].ToString().Split("/");
            DateTime dateTime = new DateTime(Convert.ToInt32(date_1[0]), Convert.ToInt32(date_1[1]), Convert.ToInt32(date_1[2]));

            string curr_Date = DateTime.Now.ToString("yyyy-MM-dd");
            string curr_Date1 = DateTime.Now.ToString("yyyy/M/dd");

            bool flag = readCSV(path + $"\\Reliability-Crashes_Date-{curr_Date}.csv", out DataTable t);

            if (flag == false)
            {
                for (int i = 0; i < 10; i++)
                {
                    flag = readCSV(path + $"\\Reliability-Crashes_Date-{DateTime.Now.AddDays(-1)}.csv", out t);
                    if (flag)
                    {
                        break;
                    }
                }
            }
            string reportDate = t.Rows[0].ItemArray[0].ToString().Split(" ")[2];//Report Date : 2020-12-05
            for (int i = 1; i < fileName.Length + 10; i++)
            {
                int lastRow1 = sheet.Range.LastRow + 1;

                string new_date1 = dateTime.AddDays(i).ToString("yyyy-MM-dd");
                string fullpath = path + Crashes_Total + new_date1 + suffix;
                string fullpath1 = path + TMAD + new_date1 + suffix;

                DateTime new_date = dateTime.AddDays(i);

                readCSV(fullpath, out DataTable dt);
                if (fileName.Contains(fullpath))
                {

                    sheet.Range["B" + lastRow1].DateTimeValue = new_date;
                    for (int x = 1; x <= name.Length; x++)
                    {
                        sheet.Range[Convert.ToChar('B' + x).ToString() + lastRow1].NumberValue = AnalyzeCSV(dt, "Crashes", name[x - 1], condition)[0];
                    }
                    sheet.Range[Convert.ToChar('B' + (name.Length + 1)).ToString() + lastRow1].NumberValue = AnalyzeCSV(dt, "Crashes", name[0], condition)[1];

                    if (fileName.Contains(fullpath1))
                    {
                        readCSV(fullpath1, out DataTable dt1);
                        if (AnalyzeCSV(dt1, "TMAD", "2004 | Vb", condition)[0] != 0)
                        {
                            sheet.Range[Convert.ToChar('B' + (name.Length + 4)).ToString() + lastRow1].NumberValue = AnalyzeCSV(dt1, "TMAD", "2004 | Vb", condition)[0];
                        }
                        else
                        {
                            sheet.Range[Convert.ToChar('B' + (name.Length + 4)).ToString() + lastRow1].NumberValue = AnalyzeCSV(dt1, "TMAD", "Insider | Vb", condition)[0];
                        }
                    }
                    else
                    {
                        int lastRow2 = lastRow1 - 1;
                        sheet.Range[Convert.ToChar('B' + (name.Length + 4)).ToString() + lastRow1].NumberValue = sheet.Range[Convert.ToChar('B' + (name.Length + 4)).ToString() + lastRow2].NumberValue;
                    }
                    sheet.Range[Convert.ToChar('B' + (name.Length + 2)).ToString() + lastRow1].Value = "||";

                    sheet.Range[Convert.ToChar('B' + (name.Length + 3)).ToString() + lastRow1].Formula = "=SUM(" + Convert.ToChar('B' + 1).ToString() + lastRow1 + ":" + Convert.ToChar('B' + (name.Length)).ToString() + lastRow1 + ")" + "/" + Convert.ToChar('B' + (name.Length + 4)).ToString() + lastRow1;

                    sheet.Range[Convert.ToChar('B' + (name.Length + 3)).ToString() + lastRow1].NumberFormat = "0.000%";


                    sheet.Range["B6:" + Convert.ToChar('B' + (name.Length + 4)).ToString() + lastRow1].HorizontalAlignment = HorizontalAlignType.Center;
                    Console.WriteLine(new_date.ToString("yyyy/M/dd"));
                }

                if (new_date.ToString("yyyy-MM-dd") == reportDate || dateTime.ToString("yyyy-MM-dd") == reportDate) { break; }
            }
            Console.WriteLine("now date :" + curr_Date);
        }
        public static void InitiExcel(string[] name)
        {
            Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
            Spire.Xls.Worksheet newSheet1 = workbook.CreateEmptySheet("Crashes");
            workbook.Worksheets.Remove("Sheet1");
            workbook.Worksheets.Remove("Sheet2");
            workbook.Worksheets.Remove("Sheet3");

            newSheet1.Range["A1:M1"].ColumnWidth = 17;
            newSheet1.Range.HorizontalAlignment = HorizontalAlignType.Center;
            newSheet1.Range["B6"].Value = "Date";

            for (int i = 1; i <= name.Length; i++)
            {
                newSheet1.Range[Convert.ToChar('B' + i).ToString() + "6"].Value = name[i - 1];
            }
            newSheet1.Range[Convert.ToChar('B' + (name.Length + 1)).ToString() + "6"].Value = "All Crashes";
            newSheet1.Range[Convert.ToChar('B' + (name.Length + 2)).ToString() + "6"].Value = "Percent Impacted";
            newSheet1.Range[Convert.ToChar('B' + (name.Length + 3)).ToString() + "6"].Value = "Crashes/OS";
            newSheet1.Range[Convert.ToChar('B' + (name.Length + 4)).ToString() + "6"].Value = "OS";
            newSheet1.Range["B6:" + Convert.ToChar('B' + (name.Length + 4)).ToString() + "6"].BorderInside(LineStyleType.Thin, Color.LightBlue);
            newSheet1.Range["B6:" + Convert.ToChar('B' + (name.Length + 4)).ToString() + "6"].BorderAround(LineStyleType.Medium, Color.LightBlue);

            workbook.SaveToFile(System.Environment.CurrentDirectory + "\\RecentData.xlsx", ExcelVersion.Version2013);
        }

        public static void GenerateChart(string[] name)
        {
            try
            {
                if (File.Exists(System.Environment.CurrentDirectory + "\\Trend Analysis.xlsx")) File.Delete(System.Environment.CurrentDirectory + "\\Trend Analysis.xlsx");
                Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
                Spire.Xls.Worksheet newSheet = workbook.CreateEmptySheet("Crashes");
                workbook.Worksheets.Remove("Sheet1");
                workbook.Worksheets.Remove("Sheet2");
                workbook.Worksheets.Remove("Sheet3");

                newSheet.Range["A1:M1"].ColumnWidth = 15;
                //newSheet.Range.HorizontalAlignment = HorizontalAlignType.Center;

                Spire.Xls.Workbook workbook1 = new Spire.Xls.Workbook();
                workbook1.LoadFromFile(System.Environment.CurrentDirectory + "\\RecentData.xlsx");
                //int lastRow = workbook1.Worksheets["Crashes"].LastRow;
                int lastRow = workbook1.Worksheets["Crashes"].LastRow;
                //workbook1.Worksheets["Crashes"].Range["B6:I6"].Copy(workbook.Worksheets["Crashes"].Range["A6:H6"]);
                //workbook1.Worksheets["Crashes"].Range["B" + (lastRow - 29) + ":I" + lastRow].Copy(workbook.Worksheets["Crashes"].Range["A7:H36"]);

                workbook1.Worksheets["Crashes"].Range["B6:" + Convert.ToChar('B' + (name.Length + 4)).ToString() + "6"].Copy(workbook.Worksheets["Crashes"].Range["A41:" + Convert.ToChar('A' + (name.Length + 4)).ToString() + "41"]);
                if (lastRow < 37)
                {
                    workbook1.Worksheets["Crashes"].Range["B7" + ":" + Convert.ToChar('B' + (name.Length + 4)).ToString() + lastRow].Copy(workbook.Worksheets["Crashes"].Range["A42:" + Convert.ToChar('A' + (name.Length + 4)).ToString() + (lastRow + 25)]);
                }
                else
                {
                    workbook1.Worksheets["Crashes"].Range["B" + (lastRow - 29) + ":" + Convert.ToChar('B' + (name.Length + 4)).ToString() + lastRow].Copy(workbook.Worksheets["Crashes"].Range["A42:" + Convert.ToChar('A' + (name.Length + 4)).ToString() + "71"]);
                }
                workbook.Worksheets["Crashes"].Range["A41:" + Convert.ToChar('A' + (name.Length + 4)).ToString() + (lastRow + 25)].HorizontalAlignment = HorizontalAlignType.Center;
                workbook.SaveToFile(System.Environment.CurrentDirectory + "\\Trend Analysis.xlsx", ExcelVersion.Version2013);
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static bool IsInt(string value)
        {
            return Regex.IsMatch(value, @"^[+-]?\d*$");
        }

        public static void Query(string Path)
        {
            bool flag = readCSV(Path, out DataTable dt);
        }
        public static int[] AnalyzeCSV(DataTable dt, string type, string Name, Dictionary<string, string[]> condition)
        {
            try
            {
                if (type == "Crashes")
                {
                    foreach (DataColumn column in dt.Columns)
                    {
                        if (column.ColumnName.Contains("["))
                        {
                            string ColumnName = column.ColumnName.Split("[")[1].Split("]")[0];
                            column.ColumnName = ColumnName.Replace(" ", "");
                        }
                    }

                    int[] crashes = { 0, 0 };
                    if (condition.ContainsKey("DriverVersion") && !condition.ContainsKey("ReleaseVersion") && !condition.ContainsKey("OSVersion"))
                    {
                        crashes[1] = dt.AsEnumerable().Select(d => Convert.ToInt32(d.Field<string>("Crashes"))).Sum();
                    }
                    else
                    {
                        crashes[1] = GetCrashes1(dt, Name, condition, "total");
                    }
                    //dt.AsEnumerable().Where(d=>d["DriverName"].ToString()== condition)
                    crashes[0] = GetCrashes1(dt, Name, condition);
                    return crashes;
                }
                else if (type == "TMAD")
                {
                    dt.Columns[0].ColumnName = "OSVersion";
                    int[] Tmad = { 0 };
                    foreach (var values in condition)
                    {
                        if (values.Key.Equals("ReleaseVersion"))
                        {
                            foreach (var item in values.Value)
                            {
                                DataRow[] dataRows = dt.Select("OSVersion = '" + item + "'");
                                Tmad[0] += Convert.ToInt32(dataRows[0].ItemArray[1]);
                            }
                            break;
                        }
                        else if (!(condition.ContainsKey("ReleaseVersion") && condition.ContainsKey("OSVersion")))
                        {
                            Tmad[0] += dt.AsEnumerable().Select(d => Convert.ToInt32(d.Field<string>("[TMAD]"))).Sum();
                            break;
                        }
                    }
                    if (condition.Count == 0)
                    {
                        Tmad[0] += dt.AsEnumerable().Select(d => Convert.ToInt32(d.Field<string>("[TMAD]"))).Sum();
                    }
                    return Tmad;
                }
                return new int[2] { 0, 0 };
            }
            catch (Exception)
            {
                throw;
            }

        }

        public static DataTable GetDT()
        {
            String path1;
            if (Directory.Exists(NetPath))
            {
                path1 = NetPath;
            }
            else
            {
                path1 = Path;
            }
            try
            {
                string[] fileName = Directory.GetFiles(path1, "*.csv");
                readCSV(fileName[^1], out DataTable dt);
                return dt;
            }
            catch (Exception)
            {
                throw;
            }
        }
        public static List<string> GetItems(DataTable dt,string name, string condition)
        {
            try
            {
                foreach (DataColumn column in dt.Columns)
                {
                    if (column.ColumnName.Contains("["))
                    {
                        string ColumnName = column.ColumnName.Split("[")[1].Split("]")[0];
                        column.ColumnName = ColumnName.Replace(" ", "");
                    }
                }
                List<string> items = dt.AsEnumerable().Where(d => d["DriverName"].ToString() == name).Select(d => d.Field<string>(condition)).Distinct().ToList<string>();
                items.Sort();
                return items;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static int GetCrashes(DataTable d, string name, Dictionary<string, string[]> condition, string flag = "crashes")
        {

            return 0;
        }
        public static int GetCrashes1(DataTable d, string name, Dictionary<string, string[]> condition, string flag = "crashes")
        {
            int crashes = 0;
            StringBuilder exp = new StringBuilder();
            string exp1 = "";
            //exp.AppendFormat(" DriverName ='{0}'", name);
            foreach (var values in condition)
            {
                if (values.Key.Equals("ReleaseVersion"))
                {
                    int i = 0;
                    exp = exp.ToString() == "" ? exp.Append("(") : exp.Append(" and (");
                    foreach (var item in values.Value)
                    {
                        if (i > 0)
                        {
                            exp.Append(" or ");
                        }
                        i++;
                        exp.AppendFormat(" ReleaseVersion ='{0}'", item);
                    }
                    exp.Append(" ) ");
                }
                if (values.Key.Equals("OSVersion"))
                {
                    int x = 0;
                    exp = exp.ToString() == "" ? exp.Append("(") : exp.Append(" and (");
                    foreach (var item in values.Value)
                    {
                        if (x > 0)
                        {
                            exp.Append(" or ");
                        }
                        x++;
                        exp.AppendFormat(" OSVersion ='{0}'", item);
                    }
                    exp.Append(" ) ");
                }
                exp1 = exp.ToString();
                if (values.Key.Equals("DriverVersion"))
                {
                    int y = 0;
                    exp = exp.ToString() == "" ? exp.Append("(") : exp.Append(" and (");
                    foreach (var item in values.Value)
                    {
                        if (y > 0)
                        {
                            exp.Append(" or ");
                        }
                        y++;
                        exp.AppendFormat(" DriverVersion ='{0}'", item);
                    }
                    exp.Append(" ) ");
                }
            }
            //string ex = " DriverName ='rltkapou64.dll' and ( ReleaseVersion = '2004 | Vb' or ReleaseVersion = '1909 | 19H2' ) ";
            try
            {
                DataRow[] dataRows = d.Select(exp.ToString());
                if (flag == "total")
                {
                    //DataRow[] dataRows1 = d.Select(exp1.ToString())
                    crashes = d.Select(exp1.ToString()).Select(d => Convert.ToInt32(d.Field<string>("Crashes"))).Sum();
                }
                else
                {
                    crashes = dataRows.Where(d => d["DriverName"].ToString() == name).Select(d => Convert.ToInt32(d.Field<string>("Crashes"))).Sum();
                }
                return crashes;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static void GenerateChart1(string[] name)
        {
            //string[] name = new string[] { "rltkapou64.dll" };
            Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
            workbook.LoadFromFile(System.Environment.CurrentDirectory + "\\Trend Analysis.xlsx");
            Spire.Xls.Worksheet sheet = workbook.Worksheets["Crashes"];
            int lastRow = sheet.Range.LastRow;
            int lastColumn = sheet.Range.LastColumn;
            int length = name.Length;
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application
            {
                Visible = true
            };
            Microsoft.Office.Interop.Excel.Workbook workbook1 = excel.Workbooks.Open(System.Environment.CurrentDirectory + "\\Trend Analysis.xlsx");
            Microsoft.Office.Interop.Excel.Chart chart = ((ChartObjects)((Microsoft.Office.Interop.Excel.Worksheet)excel.Workbooks["Trend Analysis.xlsx"].Worksheets["Crashes"]).ChartObjects()).Add(0, 0, 910, 500).Chart;

            //chart.Name = "Crashes";
            chart.HasTitle = true;
            chart.ChartTitle.Text = "DRIVER Crashes VS OS Upgrade";

            chart.ChartTitle.Select();
            //((ChartTitle)excel.Selection).Format.TextFrame2.
            chart.ChartWizard(((Microsoft.Office.Interop.Excel.Worksheet)excel.Worksheets["Crashes"]).Range["'Crashes'!A41:" + Convert.ToChar('A' + (name.Length + 1)).ToString() + lastRow + ",'Crashes'!" + Convert.ToChar('A' + (name.Length + 4)).ToString() + 41 + ":" + Convert.ToChar('A' + (name.Length + 4)).ToString() + lastRow], "63");

            chart.ChartStyle = 227;
            chart.ChartType = XlChartType.xlLineStacked;
            ((Series)chart.FullSeriesCollection("OS")).AxisGroup = XlAxisGroup.xlSecondary;
            ((Series)chart.FullSeriesCollection("All Crashes")).AxisGroup = XlAxisGroup.xlSecondary;
            ((Axis)chart.Axes("1")).MajorUnit = 4;

            ((Series)chart.FullSeriesCollection("OS")).Format.Line.ForeColor.RGB = (int)XlRgbColor.rgbBlack;
            ((Series)chart.FullSeriesCollection("All Crashes")).Format.Line.ForeColor.RGB = (int)XlRgbColor.rgbRed;
            ((Series)chart.FullSeriesCollection(name[0])).Format.Line.ForeColor.RGB = (int)XlRgbColor.rgbOrange;
            int[] ID1 = new int[(lastRow - 40) / 5];
            int[] ID2 = new int[(lastRow - 40) / 5];
            int[] ID3 = new int[(lastRow - 40) / 5];
            int[] ID4 = new int[(lastRow - 40) / 5];
            for (int i = 0; i < (lastRow - 40) / 5; i++)
            {
                ID1[i] = (i * 4) + 1;
                ID2[i] = (i * 4) + 2;
                ID3[i] = (i * 4) + 3;
                ID4[i] = (i * 4) + 4;
            }
            foreach (var item in ID1)
            {
                ((Microsoft.Office.Interop.Excel.Point)((Series)chart.FullSeriesCollection("OS")).Points(item)).HasDataLabel = true;
            }
            foreach (var item in ID2)
            {
                ((Microsoft.Office.Interop.Excel.Point)((Series)chart.FullSeriesCollection("All Crashes")).Points(item)).HasDataLabel = true;
            }
            foreach (var item in ID3)
            {
                ((Microsoft.Office.Interop.Excel.Point)((Series)chart.FullSeriesCollection(name[0])).Points(item)).HasDataLabel = true;
            }

            chart.ChartArea.Select();
            chart.SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementDataLabelCenter);
            ((Series)chart.FullSeriesCollection("OS")).Select();
            chart.SetElement(Microsoft.Office.Core.MsoChartElementType.msoElementDataLabelTop);

            workbook1.Save();
        }



        public static bool readCSV(string filePath, out DataTable dt)//从csv读取数据返回table
        {
            dt = new DataTable();
            FileStream fs = null;
            Stream rs = null;
            StreamReader sr = null;
            System.Text.Encoding encoding = Encoding.Default;
            try
            {
                if (!filePath.Contains(Environment.CurrentDirectory + "\\ExportData"))
                {
                    WebRequest req = WebRequest.Create(filePath);
                    rs = req.GetResponse().GetResponseStream();
                    sr = new System.IO.StreamReader(rs, encoding);
                }
                else
                {
                    fs = new System.IO.FileStream(filePath, System.IO.FileMode.Open,
                System.IO.FileAccess.Read);
                    sr = new System.IO.StreamReader(fs, encoding);

                }
                string strLine = "";
                string[] aryLine = null;
                string[] tableHead = null;
                int columnCount = 0;
                bool IsFirst = true;

                while ((strLine = sr.ReadLine()) != null)
                {
                    if (IsFirst == true)
                    {
                        tableHead = strLine.Split(',');
                        IsFirst = false;
                        columnCount = tableHead.Length;
                        for (int i = 0; i < columnCount; i++)
                        {
                            DataColumn dc = new DataColumn(tableHead[i]);
                            dt.Columns.Add(dc);
                        }
                    }
                    else
                    {
                        aryLine = strLine.Split(',');
                        DataRow dr = dt.NewRow();
                        for (int j = 0; j < columnCount; j++)
                        {
                            dr[j] = aryLine[j];
                        }
                        dt.Rows.Add(dr);
                    }
                }
                if (rs != null) 
                {
                    rs.Close();
                }
                if (fs != null) 
                {
                    fs.Close();
                }
                sr.Close();
                
                return true;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
            }
        }
    }
}
