﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Net;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Schema;
using Newtonsoft.Json.Linq;
using Microsoft.Office.Interop.Excel;
using System.Collections;
using System.Data;
using DataTable = System.Data.DataTable;
using System.Xml;
using Spire.Xls;
using System.Drawing;
using Spire.Xls.Charts;
using System.Text.RegularExpressions;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using Range = Microsoft.Office.Interop.Excel.Range;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using ChartTitle = Microsoft.Office.Interop.Excel.ChartTitle;
using System.Threading;
using System.Linq.Expressions;
using System.Linq.Dynamic.Core;

namespace WindowsAPI
{
    class CSharp_Win32Api
    {
        public const string NetPath = "\\\\172.30.184.28\\psd\\Common\\Auto Testing\\Auto Tools\\CrashesTool_v1.2\\ExportData\\Reliability-Crashes_Total-2020-12-12.csv";
        const string Crashes_Total = "Reliability-Crashes_Total-";
        const string TMAD = "OSAdoption-TMAD-";
        const string Crashes = "Reliability-Crashes-";

        public static string path = System.Environment.CurrentDirectory + "\\ExportData\\";
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

        const int cycleNum = 10;
        static int cnt = 10;
        static AutoResetEvent myEvent = new AutoResetEvent(false);
        // 该函数将虚拟键消息转换为字符消息。
        //[DllImport("user32.dll", CharSet = CharSet.Auto)]
        //public static extern bool TranslateMessage(ref MSG msg);


        // 该函数检取指定虚拟键的状态。
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern ushort GetKeyState(int virtKey);


        // 该函数将256个虚拟键的状态拷贝到指定的缓冲区中。
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int GetKeyboardState(byte[] pbKeyState);


        [DllImport("user32.dll")]
        public static extern void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);


        static void Main(string[] args)
        {
            string[] Name = { "rltkapou64.dll", "rltkapo64.dll", "igdkmd64.sys" };

            //ReadCSV1(out DataTable dt);
            //int count = dt.Rows.Count;
            //var dt = new DataTable();
            //readCSV("C:\\Users\\501805\\Desktop\\auto1\\TelemetryExcel1.1\\ExportData\\Reliability-Crashes_Total-09-16.csv", out dt);
            //Console.WriteLine(dt.Rows[1][1].ToString());
            //DataRow[] dataRows = dt.Select("vb");

            //analyzeCSV();
            //Console.ReadLine();
            //Progress();
            //EPPlusDemo();
            //Spire_XLSDemo();

            //GenerateExcel();

            //GenerateChart1();
            //GenerateTable();
            //Console.ReadLine();
            //InitiExcel(Name);
            Query(NetPath);

/*            ThreadPool.SetMinThreads(1, 1);
            ThreadPool.SetMaxThreads(5, 5);
            for (int i = 1; i <= cycleNum; i++)
            {
                ThreadPool.QueueUserWorkItem(new WaitCallback(testFun), i.ToString());
            }
            Console.WriteLine("主线程执行！");
            Console.WriteLine("主线程结束！");
            myEvent.WaitOne();
            Console.WriteLine("线程池终止！");
            Console.ReadKey();*/
        }
        public static void testFun(object obj)
        {
            cnt -= 1;
            Console.WriteLine(string.Format("{0}:第{1}个线程", DateTime.Now.ToString(), obj.ToString()));
            // ReadCSV1(out DataTable dt);
            Thread.Sleep(5000);
            if (cnt == 0)
            {
                myEvent.Set();
            }
        }
        public class MyDictionaryComparer : IEqualityComparer<string>
        {
            public bool Equals(string x, string y)
            {
                //throw new NotImplementedException();
                return x != y;
            }

            public int GetHashCode(string obj)
            {
                //throw new NotImplementedException();
                return obj.GetHashCode();
            }
        }
        public static Dictionary<string, string> Query(string Path)
        {

            readCSV1(Path, out DataTable dt);
            
            Dictionary<string, string> dict = new Dictionary<string, string>( )
            {
                { "name", "zhangsan" },
                { "age", "18" }
            };
            //DataColumn[] cols = new DataColumn[] { dt.Columns[2], dt.Columns[3],dt.Columns[4] };
            //dt.PrimaryKey = cols;
            //object[] objs = new object[] { "school", "class" };
            //DataRow dr = dt.Rows.Find(objs);
            bool v = dict.ContainsKey("2");
            /*            var query = from c in dt.AsEnumerable()
                                    where
                                    (String.IsNullOrEmpty(productName) || c.Field<string>("name").IndexOf(productName) > -1) &&
                                    (String.IsNullOrEmpty(CategoryID) || c.Field<string>("id").Contains(CategoryID))
                                    select c;*/
            //string ex = " DriverName ='rltkapou64.dll' and ( ReleaseVersion = '2004 | Vb' or ReleaseVersion = '1909 | 19H2' ) ";
            //DataRow[] dataRows = d.Select(exp.ToString());
            System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
            
            foreach (DataColumn column in dt.Columns)
            {
                if (column.ColumnName.Contains("["))
                {
                    string ColumnName = column.ColumnName.Split("[")[1].Split("]")[0];
                    column.ColumnName = ColumnName.Replace(" ", "");
                }
            }
            sw.Start();

            var paramExp = Expression.Variable(typeof(DataRow), "d");
            var osVersion = Expression.Constant("OSVersion", typeof(string));
            var releaseVersion = Expression.Constant("DriverVersion", typeof(string));

            var constant = Expression.Constant("10.0.19041.508");
            var constant1 = Expression.Constant("11.0.6000.614");
            Object o = "1";
            //row.Field<string>("1");
            var member = typeof(DataRowExtensions).GetMethod("Field", new Type[] { typeof(DataRow), typeof(string) }).MakeGenericMethod(typeof(string));

            MethodCallExpression methodCallExpression = Expression.Call(member, paramExp, osVersion);
            BinaryExpression binaryExpression = Expression.Equal(methodCallExpression, constant);
            BinaryExpression binaryExpression1 = Expression.Equal(Expression.Call(member, paramExp, releaseVersion), constant1);
            var expression = Expression.And(binaryExpression, binaryExpression1);
            Expression<Func<DataRow, bool>> expression1 = Expression.Lambda<Func<DataRow, bool>>(expression, paramExp);


            //BinaryExpression binaryExpression = Expression.Equal(memberExpression, constantExpression);
            //Expression<Func<DataRow, string>> expression = Expression.Lambda<Func<DataRow, string>>(binaryExpression, paramExp);

            EnumerableRowCollection<DataRow> enumerableRowCollections = from c in dt.AsEnumerable()
                                                                        where
                                                                        c.Field<string>("OSVersion").Equals("10.0.19041.508") &&
                                                                        c.Field<string>("ReleaseVersion").Equals("2004 | Vb") || c.Field<string>("ReleaseVersion").Equals("1909 | 19H2") &&
                                                                        c.Field<string>("DriverVersion").Equals("11.0.6000.614")
                                                                        select c;
            int v1 = dt.AsEnumerable().AsQueryable().Where(expression1).Select(d => Convert.ToInt32(d.Field<string>("Crashes"))).Sum();
            dt.AsEnumerable().AsQueryable().Where(d => d["OSVersion"].ToString() == "10.0.19041.508");
            Console.WriteLine(v1);
            /*            string ex = " OSVersion ='10.0.19041.508' and ( ReleaseVersion = '2004 | Vb' or ReleaseVersion = '1909 | 19H2' )and(DriverVersion = '11.0.6000.614') ";
                        DataRow[] dataRows = dt.Select(ex);*/
            int crashes = 0,total = 0;
            //crashes = dt.Select(ex.ToString()).Select(d => Convert.ToInt32(d.Field<string>("Crashes"))).Sum();
            //total = dataRows.Where(d => d["DriverName"].ToString() == "rltkapo64.dll").Select(d => Convert.ToInt32(d.Field<string>("Crashes"))).Sum();
            total = enumerableRowCollections.Select(d => Convert.ToInt32(d.Field<string>("Crashes"))).Sum();
            crashes = enumerableRowCollections.Where(d => d["DriverName"].ToString() == "rltkapo64.dll").Select(d => Convert.ToInt32(d.Field<string>("Crashes"))).Sum();
            sw.Stop();
            TimeSpan ts2 = sw.Elapsed;
            Console.WriteLine("example2 time {0} ms", ts2.TotalMilliseconds);
            return dict;
        }
        public static bool ReadCSV1(out DataTable dt)//从csv读取数据返回table
        {
            dt = new DataTable();
            try
            {
                System.Text.Encoding encoding = Encoding.Default;
                WebRequest req = WebRequest.Create(NetPath);
                Stream rs = req.GetResponse().GetResponseStream();
                StreamReader sr = new StreamReader(rs, encoding);

                //记录每次读取的一行记录
                string strLine = "";
                //记录每行记录中的各字段内容
                string[] aryLine = null;
                string[] tableHead = null;
                int columnCount = 0;
                //标示是否是读取的第一行
                bool IsFirst = true;
                //逐行读取CSV中的数据
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
                if (aryLine != null && aryLine.Length > 0)
                {
                    dt.DefaultView.Sort = tableHead[0] + " " + "asc";
                }
                sr.Close();
                rs.Close();
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        public static void test()
        {
            
        }
        public static void Progress()
        {
            Console.WriteLine("-------Beginning Working -------");

            Console.WriteLine(" 0 %");
            for (int i = 0; ++i <= 100;)
            {
                Console.SetCursorPosition(1, 1);
                Console.Write(" {0} %", i);
                //模拟实际工作中的延迟,否则进度太快
                System.Threading.Thread.Sleep(100);
            }
        }

        public static void TestProgress()
        {
            bool isBreak = false;
            ConsoleColor colorBack = Console.BackgroundColor;
            ConsoleColor colorFore = Console.ForegroundColor;

            //第一行信息
            Console.WriteLine("****** now working...******");

            //第二行绘制进度条背景
            Console.BackgroundColor = ConsoleColor.DarkCyan;
            for (int i = 0; ++i <= 25;)
            {
                Console.Write(" ");
            }
            Console.WriteLine(" ");
            Console.BackgroundColor = colorBack;

            //第三行输出进度
            Console.WriteLine("0%");
            //第四行输出提示,按下回车可以取消当前进度
            Console.WriteLine("<Press Enter To Break.>");
            //以上绘制一个完整的工作区域

            //开始控制进度条和进度变化
            for (int i = 0; ++i <= 100;)
            {
                //先检查是否有按键请求,如果有,判断是否为回车键,如果是则退出循环
                if (Console.KeyAvailable && System.Console.ReadKey(true).Key == ConsoleKey.Enter)
                {
                    isBreak = true; break;
                }
                //绘制进度条进度
                Console.BackgroundColor = ConsoleColor.Yellow;//设置进度条颜色
                Console.SetCursorPosition(i / 4, 1);//设置光标位置,参数为第几列和第几行
                Console.Write(" ");//移动进度条
                Console.BackgroundColor = colorBack;//恢复输出颜色
                                                    //更新进度百分比,原理同上.
                Console.ForegroundColor = ConsoleColor.Green;
                Console.SetCursorPosition(0, 2);
                Console.Write("{0}%", i);
                Console.ForegroundColor = colorFore;
                //模拟实际工作中的延迟,否则进度太快
                System.Threading.Thread.Sleep(100);
            }
            //工作完成,根据实际情况输出信息,而且清楚提示退出的信息
            Console.SetCursorPosition(0, 3);
            Console.Write(isBreak ? "break!!!" : "finished.");
            Console.WriteLine(" ");
            //等待退出
            Console.ReadKey(true);
        }
        public static void GenerateExcel()
        {
            Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
            //Initailize worksheet
            //workbook.CreateEmptySheets(1);
            if (File.Exists("RecentData.xlsx"))
            {
                workbook.LoadFromFile("RecentData.xlsx");
            }
            else
            {
                InitiExcel(Name);
                workbook.LoadFromFile("RecentData.xlsx");
            }
            //LoadData(workbook, "Crashes",Name ,Condition);
            LoadData(workbook, "Total_Crashes", Name, Condition);

            workbook.SaveToFile("RecentData.xlsx", ExcelVersion.Version2013);

            
        }
        public static void LoadData(Spire.Xls.Workbook workbook, string sheetName, string[] name, Dictionary<string, string> Condition)
        {
            Spire.Xls.Worksheet sheet = workbook.Worksheets[sheetName];
            int lastRow = sheet.Range.LastRow;              //2020/10/12 0:00:00

            string[] date_1 = sheet.Range["B" + lastRow].Value == "Date" ? "2020/7/1".Split("/") : sheet.Range["B" + lastRow].Value.Split(" ")[0].ToString().Split("/");
            DateTime dateTime = new DateTime(Convert.ToInt32(date_1[0]), Convert.ToInt32(date_1[1]), Convert.ToInt32(date_1[2]));

            string curr_Date = DateTime.Now.ToString("MM-dd");
            string curr_Date1 = DateTime.Now.ToString("yyyy/M/dd");
            for (int i = 1; i < 1000; i++)
            {
                string[] fileName = Directory.GetFiles("C:\\Users\\501805\\Desktop\\auto1\\TelemetryExcel1.1\\ExportData", "*.csv");

                int lastRow1 = sheet.Range.LastRow + 1;

                string new_date1 = dateTime.AddDays(i).ToString("MM-dd");
                string fullpath = path + Crashes_Total + new_date1 + suffix;
                string fullpath1 = path + TMAD + new_date1 + suffix;
                DateTime new_date = dateTime.AddDays(i);
                if (fileName.Contains(fullpath))
                {
                    sheet.Range["B" + lastRow1].DateTimeValue = new_date;
                    for (int x = 1; x <= name.Length; x++)
                    {
                        sheet.Range[Convert.ToChar('B' + x).ToString() + lastRow1].NumberValue = AnalyzeCSV(fullpath, "Crashes", name[x - 1], Condition)[0];
                    }
                    sheet.Range[Convert.ToChar('B' + (name.Length + 1)).ToString() + lastRow1].NumberValue = AnalyzeCSV(fullpath, "Crashes", name[0], Condition)[1];

                    if (fileName.Contains(fullpath1))
                    {
                        if (AnalyzeCSV(fullpath1, "TMAD", "2004 | Vb", Condition)[0] != 0)
                        {
                            sheet.Range[Convert.ToChar('B' + (name.Length + 4)).ToString() + lastRow1].NumberValue = AnalyzeCSV(fullpath1, "TMAD", "2004 | Vb", Condition)[0];
                        }
                        else
                        {
                            sheet.Range[Convert.ToChar('B' + (name.Length + 4)).ToString() + lastRow1].NumberValue = AnalyzeCSV(fullpath1, "TMAD", "Insider | Vb", Condition)[0];
                        }
                    }
                    else
                    {
                        int lastRow2 = lastRow1 - 1;
                        sheet.Range[Convert.ToChar('B' + (name.Length + 4)).ToString() + lastRow1].NumberValue = sheet.Range["I" + lastRow2].NumberValue;
                    }
                    sheet.Range[Convert.ToChar('B' + (name.Length + 3)).ToString() + lastRow1].Formula = "=(C" + lastRow1 + "+D" + lastRow1 + "+E" + lastRow1 + ")/I" + lastRow1;
                    sheet.Range[Convert.ToChar('B' + (name.Length + 3)).ToString() + lastRow1].NumberFormat = "0.000%";

                    sheet.Range[Convert.ToChar('B' + (name.Length + 2)).ToString() + lastRow1].Value = "||";

                    Console.WriteLine(new_date.ToString("yyyy/M/dd"));
                }

                if (new_date.ToString("yyyy/M/dd") == curr_Date1 || dateTime.ToString("yyyy/M/dd") == curr_Date1) { break; }
            }
            Console.WriteLine("now date :" + curr_Date);
        }
        public static void InitiExcel(string[] name)
        {
            Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
            Spire.Xls.Worksheet newSheet = workbook.CreateEmptySheet("Crashes");
            workbook.Worksheets.Remove("sheet1");
            workbook.Worksheets.Remove("sheet2");
            workbook.Worksheets.Remove("sheet3");

            newSheet.Range["A1:M1"].ColumnWidth = 15;
            //newSheet.Range["A:L"].HorizontalAlignment = HorizontalAlignType.Center;
            newSheet.Range["B6"].Value = "Date";
            for (int i = 1; i <= name.Length; i++)
            {
                newSheet.Range[Convert.ToChar('B' + i).ToString() + "6"].Value = name[i - 1];
            }
            newSheet.Range[Convert.ToChar('B' + (name.Length + 1)).ToString() + "6"].Value = "All Crashes";
            newSheet.Range[Convert.ToChar('B' + (name.Length + 2)).ToString() + "6"].Value = "Percent Impacted";
            newSheet.Range[Convert.ToChar('B' + (name.Length + 3)).ToString() + "6"].Value = "Crashes/20H1";
            newSheet.Range[Convert.ToChar('B' + (name.Length + 4)).ToString() + "6"].Value = "20H1";
            newSheet.Range["B6:" + Convert.ToChar('B' + (name.Length + 4)).ToString() + "6"].BorderInside(LineStyleType.Thin, Color.LightBlue);
            newSheet.Range["B6:" + Convert.ToChar('B' + (name.Length + 4)).ToString() + "6"].BorderAround(LineStyleType.Medium, Color.LightBlue);

            Spire.Xls.Worksheet newSheet1 = workbook.CreateEmptySheet("Total_Crashes");
            newSheet1.Range["A1:M1"].ColumnWidth = 15;
            newSheet1.Range.HorizontalAlignment = HorizontalAlignType.Center;
            newSheet1.Range["B6"].Value = "Date";

            for (int i = 1; i <= name.Length; i++)
            {
                newSheet1.Range[Convert.ToChar('B' + i).ToString() + "6"].Value = name[i - 1];
            }
            newSheet1.Range[Convert.ToChar('B' + (name.Length + 1)).ToString() + "6"].Value = "All Crashes";
            newSheet1.Range[Convert.ToChar('B' + (name.Length + 2)).ToString() + "6"].Value = "Percent Impacted";
            newSheet1.Range[Convert.ToChar('B' + (name.Length + 3)).ToString() + "6"].Value = "Crashes/20H1";
            newSheet1.Range[Convert.ToChar('B' + (name.Length + 4)).ToString() + "6"].Value = "20H1";
            newSheet1.Range["B6:" + Convert.ToChar('B' + (name.Length + 4)).ToString() + "6"].BorderInside(LineStyleType.Thin, Color.LightBlue);
            newSheet1.Range["B6:" + Convert.ToChar('B' + (name.Length + 4)).ToString() + "6"].BorderAround(LineStyleType.Medium, Color.LightBlue);

            workbook.SaveToFile("RecentData.xlsx", ExcelVersion.Version2013);
        }

        public static void GenerateChart()
        {
            if (File.Exists("Trend Analysis.xlsx")) File.Delete("Trend Analysis.xlsx");
            Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
            Spire.Xls.Worksheet newSheet = workbook.CreateEmptySheet("Crashes");
            Spire.Xls.Worksheet newSheet1 = workbook.CreateEmptySheet("Total_Crashes");
            workbook.Worksheets.Remove("sheet1");
            workbook.Worksheets.Remove("sheet2");
            workbook.Worksheets.Remove("sheet3");

            newSheet.Range["A1:M1"].ColumnWidth = 15;
            newSheet1.Range["A1:M1"].ColumnWidth = 15;
            //newSheet.Range.HorizontalAlignment = HorizontalAlignType.Center;


            Spire.Xls.Workbook workbook1 = new Spire.Xls.Workbook();
            workbook1.LoadFromFile("RecentData.xlsx");
            int lastRow = workbook1.Worksheets["Crashes"].LastRow;
            int lastRow1 = workbook1.Worksheets["Total_Crashes"].LastRow;
            workbook1.Worksheets["Crashes"].Range["B6:I6"].Copy(workbook.Worksheets["Crashes"].Range["A6:H6"]);
            workbook1.Worksheets["Crashes"].Range["B" + (lastRow - 29) + ":I" + lastRow].Copy(workbook.Worksheets["Crashes"].Range["A7:H36"]);

            workbook1.Worksheets["Total_Crashes"].Range["B6:I6"].Copy(workbook.Worksheets["Total_Crashes"].Range["A6:H6"]);
            workbook1.Worksheets["Total_Crashes"].Range["B" + (lastRow1 - 29) + ":I" + lastRow1].Copy(workbook.Worksheets["Total_Crashes"].Range["A7:H36"]);
            workbook.SaveToFile("Trend Analysis.xlsx", ExcelVersion.Version2013);


        }

        public static bool IsInt(string value)
        {
            return Regex.IsMatch(value, @"^[+-]?\d*$");
        }
        public static int[] AnalyzeCSV(string path, string type, string Name, Dictionary<string, string> condition)
        {
            if (type == "Crashes")
            {
                readCSV(path, out DataTable dt);
                foreach (DataColumn column in dt.Columns)
                {
                    if (column.ColumnName.Contains("["))
                    {
                        string ColumnName = column.ColumnName.Split("[")[1].Split("]")[0];
                        column.ColumnName = ColumnName.Replace(" ", "");
                    }
                }
                int[] crashes = { 0, 0 };
                //total = dt.AsEnumerable()
                //                   .Where(r => r.Field<string>("Crashes") != null)
                //                   .Sum(x => Convert.ToInt32(x.Field<string>("Crashes")));
                //object total = dt.Compute("SUM(Crashes)", "");
                double total = 0;
                foreach (DataRow row in dt.Rows)
                {
                    if (IsInt(row["Crashes"].ToString()))
                    {
                        total += double.Parse(row["Crashes"].ToString());
                    }
                    else
                    {
                        continue;
                    }
                }
                crashes[1] = Convert.ToInt32(total);
                //dt.AsEnumerable().Where(r => r.Field<string>("OSVersion") != null).Select("DriverName= '" + condition + "' AND ISNULL(OSVersion = '10.0.19042.572',0)");
                //   AND ISNULL(DriverVersion = '26.20.100.7872',0)',0)AND ISNULL(OSVersion = '26.20.100.7872',0)AND ISNULL(DriverVersion = '26.20.100.7872',0)
                //(DataRow i in dt.Select("DriverName= '" + condition + "' AND ISNULL(DriverVersion = '11.0.6000.620',0) AND ISNULL(OSVersion = '10.0.19042.572',0) "))

                //dt.AsEnumerable().Where(d=>d["DriverName"].ToString()== condition)
                foreach (DataRow i in editExp(dt, Name, condition)) //"DriverName = 'rltkapou64.dll'"
                {
                    object crash = i["Crashes"];
                    crashes[0] += Convert.ToInt32(crash);
                }
                //dt.AsEnumerable().Select(d => d.Field<string>("OSVersion")).Distinct().ToList<string>()
                return crashes;
            }
            else if (type == "TMAD")
            {
                bool flag = readCSV(path, out DataTable dt);

                dt.Columns[0].ColumnName = "OSVersion";
                int[] Tmad = { 0 };
                foreach (DataRow i in dt.Select("OSVersion= '" + condition + "'"))  //"OSVersion= '2004 | Vb'"
                {
                    object tmad = i.ItemArray[1];
                    Tmad[0] += Convert.ToInt32(tmad);
                }
                return Tmad;
            }
            return new int[2] { 0, 0 };
        }

        public static EnumerableRowCollection<DataRow> editExp(DataTable d, string name, Dictionary<string, string> condition)
        {

            EnumerableRowCollection<DataRow> enumerable = d.AsEnumerable().Where(d => d["DriverName"].ToString() == name).Where(d => d["DriverVersion"].ToString() == "11.0.6000.620").Where(d => d["OSVersion"].ToString() == "10.0.19042.572");

            return enumerable;
        }

        static public bool readCSV(string filePath, out DataTable dt)//从csv读取数据返回table
        {
            dt = new DataTable();
            try
            {
                System.Text.Encoding encoding = Encoding.Default;//GetType(filePath); //
                                                                 // DataTable dt = new DataTable();
                System.IO.FileStream fs = new System.IO.FileStream(filePath, System.IO.FileMode.Open,
                    System.IO.FileAccess.Read);
                System.IO.StreamReader sr = new System.IO.StreamReader(fs, encoding);
                //记录每次读取的一行记录
                string strLine = "";
                //记录每行记录中的各字段内容
                string[] aryLine = null;
                string[] tableHead = null;
                int columnCount = 0;
                //标示是否是读取的第一行
                bool IsFirst = true;
                //逐行读取CSV中的数据
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
                if (aryLine != null && aryLine.Length > 0)
                {
                    dt.DefaultView.Sort = tableHead[0] + " " + "asc";
                }
                sr.Close();
                fs.Close();
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static void GenerateTable()
        {
            if (File.Exists("Trend Analysis.xlsx")) File.Delete("Trend Analysis.xlsx");
            Application excel = new Application();
            excel.Visible = true;
            excel.Workbooks.Open("C:\\Users\\501805\\source\\repos\\Test\\bin\\Debug\\netcoreapp3.1\\RecentData.xlsx");

            Worksheet worksheet = (Worksheet)excel.Sheets["Crashes"];
            int CrashesRow3 = worksheet.Range["B65535"].End[XlDirection.xlUp].Row;
            Application excel1 = New_Excel();

            ((Worksheet)excel.Workbooks["RecentData.xlsx"].Worksheets["Crashes"]).Range["B6:I6"].Copy();
            Thread.Sleep(1000);
            ((Worksheet)excel1.Worksheets["Trend Analysis"]).Range["A6"].PasteSpecial(XlPasteType.xlPasteValues);

            ((Worksheet)excel.Workbooks["RecentData.xlsx"].Worksheets["Crashes"]).Range["B"+(CrashesRow3 - 29) + ":I" +CrashesRow3].Copy();
            Thread.Sleep(1000);
            ((Worksheet)excel1.Sheets["Trend Analysis"]).Range["A7:H36"].PasteSpecial(XlPasteType.xlPasteValues);
            ((Worksheet)excel1.Sheets["Trend Analysis"]).ListObjects.Add(XlListObjectSourceType.xlSrcRange,excel1.Range["A6:H36"]);

            //((Worksheet)excel1.Sheets["Trend Analysis"]).ListObjects[1].TableStyle = TableBuiltInStyles.TableStyleDark10;


            excel.Workbooks["RecentData.xlsx"].Save();
            excel.Windows["RecentData.xlsx"].Close(false);  


            
        }
        public static void GenerateChart1()
        {
            Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
            workbook.LoadFromFile(System.Environment.CurrentDirectory + "\\Trend Analysis.xlsx");
            Spire.Xls.Worksheet sheet = workbook.Worksheets["Crashes"];
            int lastRow = sheet.Range.LastRow;
            int lastColumn = sheet.Range.LastColumn;
            string[] name = new string[] { "rltkapou64.dll" };
            int length = name.Length;
            Application excel = new Application
            {
                Visible = true
            };
            excel.Workbooks.Open("C:\\Users\\501805\\source\\repos\\Test\\bin\\Debug\\netcoreapp3.1\\Trend Analysis.xlsx");
            
            //((ChartObjects)((Worksheet)excel.Workbooks["Trend Analysis.xlsx"].Worksheets["Trend Analysis"]).ChartObjects()).Add(60, 570, 950, 500);
            //Console.WriteLine(count);
            Microsoft.Office.Interop.Excel.Chart chart = ((ChartObjects)((Worksheet)excel.Workbooks["Trend Analysis.xlsx"].Worksheets["Crashes"]).ChartObjects()).Add(20, 570, 910, 500).Chart;

            //chart.Name = "Crashes";
            chart.HasTitle = true;
            chart.ChartTitle.Text = "DRIVER Crashes VS OS Upgrade";

            chart.ChartTitle.Select();
            string v = "'Crashes'!A6:" + Convert.ToChar('A' + (name.Length + 1)).ToString() + lastRow + ",'Crashes'!" + Convert.ToChar('A' + (name.Length + 1)).ToString() + 6 + ":" + Convert.ToChar('A' + (name.Length + 4)).ToString() + lastRow;
            //((ChartTitle)excel.Selection).Format.TextFrame2.
            chart.ChartWizard(((Worksheet)excel.Worksheets["Crashes"]).Range["'Crashes'!A6:"+ Convert.ToChar('A' + (name.Length + 1)).ToString()+lastRow+",'Crashes'!"+ Convert.ToChar('A' + (name.Length + 4)).ToString()+6+":"+ Convert.ToChar('A' + (name.Length + 4)).ToString()+lastRow], "63");

            //chart.ChartWizard(((Worksheet)excel.Worksheets["Crashes"]).Range["'Crashes'!A6:E36,'Crashes'!H6:H36"], "63");
            //.SetSourceData(Range("表1[[#All],[20H1]]"))
            //chart.Tab.ThemeColor = XlThemeColor.xlThemeColorAccent6;
            chart.ChartStyle = 227;
            chart.ChartType = XlChartType.xlLineStacked;
            ((Series)chart.FullSeriesCollection("OS")).AxisGroup = XlAxisGroup.xlSecondary;
            ((Series)chart.FullSeriesCollection("All Crashes")).AxisGroup = XlAxisGroup.xlSecondary;
            ((Axis)chart.Axes("1")).MajorUnit = 4;

            ((Series)chart.FullSeriesCollection("OS")).Format.Line.ForeColor.RGB = (int)XlRgbColor.rgbBlack;
            ((Series)chart.FullSeriesCollection("All Crashes")).Format.Line.ForeColor.RGB = (int)XlRgbColor.rgbRed;
            ((Series)chart.FullSeriesCollection(name[0])).Format.Line.ForeColor.RGB = (int)XlRgbColor.rgbOrange;
            int[] ID1 = new int[lastRow / 5];
            int[] ID2 = new int[lastRow / 5];
            int[] ID3 = new int[lastRow / 5];
            for (int i = 0; i < lastRow/5; i++)
            {
                ID1[i] = (i * 4) + 1;
                ID2[i] = (i * 4) + 2;
                ID3[i] = (i * 4) + 3;
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

        }


        public static Application New_Excel()
        {
            Application excel = new Application
            {
                Visible = true
            };
            Workbook workbook = excel.Workbooks.Add();
            excel.WindowState = XlWindowState.xlMaximized;
            ((Worksheet)excel.Sheets["Sheet1"]).Activate();
            string ControlFile = excel.ActiveWorkbook.Name;
            ((Worksheet)excel.ActiveSheet).Name = "Trend Analysis";
            excel.ActiveWindow.Zoom = 80; 
            excel.Cells.HorizontalAlignment = -4108;
            excel.Cells.Font.Name = "Calibri";
            excel.Cells.Interior.PatternColorIndex = -4105; 
            ((Range)excel.Columns["A:O"]).ColumnWidth = 19.27;

            ((Worksheet)excel.Workbooks[ControlFile].Sheets["Trend Analysis"]).Activate();
            excel.Workbooks[ControlFile].SaveAs("C:\\Users\\501805\\source\\repos\\Test\\bin\\Debug\\netcoreapp3.1\\Trend Analysis.xlsx");
            return excel;
        }

        public static void Spire_XLSDemo()
        {
            Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();

            //Initailize worksheet
            workbook.CreateEmptySheets(1);
            Spire.Xls.Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Chart data";
            sheet.GridLinesVisible = false;

            //Writes chart data
            CreateChartData(sheet);
            //Add a new  chart worsheet to workbook
            Spire.Xls.Chart chart = sheet.Charts.Add();
            chart.ChartType = ExcelChartType.Line;
            //Set region of chart data
            chart.DataRange = sheet.Range["A1:E5"];

            
            //Set position of chart
            chart.LeftColumn = 1;
            chart.TopRow = 6;
            chart.RightColumn = 11;
            chart.BottomRow = 29;


            //Chart title
            chart.ChartTitle = "Sales market by country";
            chart.ChartTitleArea.IsBold = true;
            chart.ChartTitleArea.Size = 12;

            chart.PrimaryCategoryAxis.Title = "Month";
            chart.PrimaryCategoryAxis.Font.IsBold = true;
            chart.PrimaryCategoryAxis.TitleArea.IsBold = true;

            chart.PrimaryValueAxis.HasMajorGridLines = false;
            chart.PrimaryValueAxis.TitleArea.TextRotationAngle = 90;
            chart.PrimaryValueAxis.MinValue = 1000;
            chart.PrimaryValueAxis.TitleArea.IsBold = true;

            
            Spire.Xls.Charts.ChartArea chartArea = chart.ChartArea;

            chart.Legend.Position = LegendPositionType.Corner;
            

            foreach (ChartSerie cs in chart.Series)
            {
                cs.Format.Options.IsVaryColor = true;
                cs.DataPoints.DefaultDataPoint.DataLabels.HasValue = true;

            }
            
            chart.PlotArea.Fill.Visible = false;

            chart.Legend.Position = LegendPositionType.Top;
            workbook.SaveToFile("Sample.xlsx",ExcelVersion.Version2016);

        }
        private static void CreateChartData(Spire.Xls.Worksheet sheet)
        {
            //Country
            sheet.Range["A1"].Value = "Country";
            sheet.Range["A2"].Value = "Cuba";
            sheet.Range["A3"].Value = "Mexico";
            sheet.Range["A4"].Value = "France";
            sheet.Range["A5"].Value = "German";

            //Jun
            sheet.Range["B1"].Value = "Jun";
            sheet.Range["B2"].NumberValue = 3300;
            sheet.Range["B3"].NumberValue = 2300;
            sheet.Range["B4"].NumberValue = 4500;
            sheet.Range["B5"].NumberValue = 6700;

            //Jul
            sheet.Range["C1"].Value = "Jul";
            sheet.Range["C2"].NumberValue = 7500;
            sheet.Range["C3"].NumberValue = 2900;
            sheet.Range["C4"].NumberValue = 2300;
            sheet.Range["C5"].NumberValue = 4200;

            //Aug
            sheet.Range["D1"].Value = "Aug";
            sheet.Range["D2"].NumberValue = 7700;
            sheet.Range["D3"].NumberValue = 6900;
            sheet.Range["D4"].NumberValue = 8400;
            sheet.Range["D5"].NumberValue = 4200;

            //Sep
            sheet.Range["E1"].Value = "Sep";
            sheet.Range["E2"].NumberValue = 8000;
            sheet.Range["E3"].NumberValue = 7200;
            sheet.Range["E4"].NumberValue = 8100;
            sheet.Range["E5"].NumberValue = 5600;

            //Style
            sheet.Range["A1:E1"].Style.Font.IsBold = true;
            sheet.Range["A2:E2"].Style.KnownColor = ExcelColors.LightYellow;
            sheet.Range["A3:E3"].Style.KnownColor = ExcelColors.LightGreen1;
            sheet.Range["A4:E4"].Style.KnownColor = ExcelColors.LightOrange;
            sheet.Range["A5:E5"].Style.KnownColor = ExcelColors.LightTurquoise;

            //Border
            sheet.Range["A1:E5"].Style.Borders[BordersLineType.EdgeTop].Color = Color.FromArgb(0, 0, 128);
            sheet.Range["A1:E5"].Style.Borders[BordersLineType.EdgeTop].LineStyle = LineStyleType.Thin;
            sheet.Range["A1:E5"].Style.Borders[BordersLineType.EdgeBottom].Color = Color.FromArgb(0, 0, 128);
            sheet.Range["A1:E5"].Style.Borders[BordersLineType.EdgeBottom].LineStyle = LineStyleType.Thin;
            sheet.Range["A1:E5"].Style.Borders[BordersLineType.EdgeLeft].Color = Color.FromArgb(0, 0, 128);
            sheet.Range["A1:E5"].Style.Borders[BordersLineType.EdgeLeft].LineStyle = LineStyleType.Thin;
            sheet.Range["A1:E5"].Style.Borders[BordersLineType.EdgeRight].Color = Color.FromArgb(0, 0, 128);
            sheet.Range["A1:E5"].Style.Borders[BordersLineType.EdgeRight].LineStyle = LineStyleType.Thin;

            sheet.Range["B2:D5"].Style.NumberFormat = "\"$\"#,##0";

            
        }
        //    private void btn_NPOI_Click(object sender, EventArgs e)
        //{
        //    string importExcelPath = "E:\\import.xlsx";
        //    string exportExcelPath = "E:\\export.xlsx";
        //    NPOI.SS.UserModel.IWorkbook workbook = WorkbookFactory.Create(importExcelPath);
        //    ISheet sheet = workbook.GetSheetAt(0);//获取第一个工作薄
        //    IRow row = (IRow)sheet.GetRow(0);//获取第一行

        //    //设置第一行第一列值,更多方法请参考源官方Demo
        //    row.CreateCell(0).SetCellValue("test");//设置第一行第一列值

        //    //导出excel
        //    FileStream fs = new FileStream(exportExcelPath, FileMode.Create, FileAccess.ReadWrite);
        //    workbook.Write(fs);
        //    fs.Close();
        //}


        
        public static string HttpGet(string Url)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(Url);
            request.Method = "GET";
            request.ContentType = "text/html;charset=UTF-8";
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            Stream myResponseStream = response.GetResponseStream();
            StreamReader myStreamReader = new StreamReader(myResponseStream, Encoding.GetEncoding("utf-8"));
            string retString = myStreamReader.ReadToEnd();
            myStreamReader.Close();
            myResponseStream.Close();
            // Console.WriteLine(retString);
            return retString;
        }

        public void TestFunc()
        {
            //keybd_event(255, 0, 0, 0);
            keybd_event(144, 0, 0, 0);
            //keybd_event(255, 0, 0x0002, 0);
            keybd_event(144, 0, 0x0002, 0);
            ushort v = GetKeyState(0x79);
            Console.WriteLine(v);
        }
        public static DataTable OpenCSV(string filePath)
        {
            Encoding encoding = GetType(filePath); //Encoding.ASCII;//
            DataTable dt = new DataTable();
            FileStream fs = new FileStream(filePath, System.IO.FileMode.Open, System.IO.FileAccess.Read);

            //StreamReader sr = new StreamReader(fs, Encoding.UTF8);
            StreamReader sr = new StreamReader(fs, encoding);
            //string fileContent = sr.ReadToEnd();
            //encoding = sr.CurrentEncoding;
            //记录每次读取的一行记录
            string strLine = "";
            //记录每行记录中的各字段内容
            string[] aryLine = null;
            string[] tableHead = null;
            //标示列数
            int columnCount = 0;
            //标示是否是读取的第一行
            bool IsFirst = true;
            //逐行读取CSV中的数据
            while ((strLine = sr.ReadLine()) != null)
            {
                //strLine = Common.ConvertStringUTF8(strLine, encoding);
                //strLine = Common.ConvertStringUTF8(strLine);

                if (IsFirst == true)
                {
                    tableHead = strLine.Split(',');
                    IsFirst = false;
                    columnCount = tableHead.Length;
                    //创建列
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
            if (aryLine != null && aryLine.Length > 0)
            {
                dt.DefaultView.Sort = tableHead[0] + " " + "asc";
            }

            sr.Close();
            fs.Close();
            return dt;
        }

        public static System.Text.Encoding GetType(string FILE_NAME)
        {
            System.IO.FileStream fs = new System.IO.FileStream(FILE_NAME, System.IO.FileMode.Open,
                System.IO.FileAccess.Read);
            System.Text.Encoding r = GetType(fs);
            fs.Close();
            return r;
        }

        /// 通过给定的文件流，判断文件的编码类型
        /// <param name="fs">文件流</param>
        /// <returns>文件的编码类型</returns>
        public static System.Text.Encoding GetType(System.IO.FileStream fs)
        {
            byte[] Unicode = new byte[] { 0xFF, 0xFE, 0x41 };
            byte[] UnicodeBIG = new byte[] { 0xFE, 0xFF, 0x00 };
            byte[] UTF8 = new byte[] { 0xEF, 0xBB, 0xBF }; //带BOM
            System.Text.Encoding reVal = System.Text.Encoding.Default;

            System.IO.BinaryReader r = new System.IO.BinaryReader(fs, System.Text.Encoding.Default);
            int i;
            int.TryParse(fs.Length.ToString(), out i);
            byte[] ss = r.ReadBytes(i);
            if (IsUTF8Bytes(ss) || (ss[0] == 0xEF && ss[1] == 0xBB && ss[2] == 0xBF))
            {
                reVal = System.Text.Encoding.UTF8;
            }
            else if (ss[0] == 0xFE && ss[1] == 0xFF && ss[2] == 0x00)
            {
                reVal = System.Text.Encoding.BigEndianUnicode;
            }
            else if (ss[0] == 0xFF && ss[1] == 0xFE && ss[2] == 0x41)
            {
                reVal = System.Text.Encoding.Unicode;
            }
            r.Close();
            return reVal;
        }

        /// 判断是否是不带 BOM 的 UTF8 格式
        /// <param name="data"></param>
        /// <returns></returns>
        private static bool IsUTF8Bytes(byte[] data)
        {
            int charByteCounter = 1;  //计算当前正分析的字符应还有的字节数
            byte curByte; //当前分析的字节.
            for (int i = 0; i < data.Length; i++)
            {
                curByte = data[i];
                if (charByteCounter == 1)
                {
                    if (curByte >= 0x80)
                    {
                        //判断当前
                        while (((curByte <<= 1) & 0x80) != 0)
                        {
                            charByteCounter++;
                        }
                        //标记位首位若为非0 则至少以2个1开始 如:110XXXXX...........1111110X　
                        if (charByteCounter == 1 || charByteCounter > 6)
                        {
                            return false;
                        }
                    }
                }
                else
                {
                    //若是UTF-8 此时第一位必须为1
                    if ((curByte & 0xC0) != 0x80)
                    {
                        return false;
                    }
                    charByteCounter--;
                }
            }
            if (charByteCounter > 1)
            {
                throw new Exception("非预期的byte格式");
            }
            return true;
        }

        public static void GenerateExcel(string[] name, Dictionary<string, string[]> condition)
        {
            Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
            //Initailize worksheet
            //workbook.CreateEmptySheets(1);
            if (File.Exists(System.Environment.CurrentDirectory + "\\RecentData.xlsx"))
            {
                workbook.LoadFromFile(System.Environment.CurrentDirectory + "\\RecentData.xlsx");
            }
            else
            {
                InitiExcel(name);
                workbook.LoadFromFile(System.Environment.CurrentDirectory + "\\RecentData.xlsx");
            }
            //LoadData(workbook, "Crashes",Name ,Condition);
            LoadData(workbook, "Crashes", name, condition);

            workbook.SaveToFile(System.Environment.CurrentDirectory + "\\RecentData.xlsx", ExcelVersion.Version2013);

            GenerateChart(name);
        }
        public static void LoadData(Spire.Xls.Workbook workbook, string sheetName, string[] name, Dictionary<string, string[]> condition)
        {
            Spire.Xls.Worksheet sheet = workbook.Worksheets[sheetName];
            int lastRow = sheet.Range.LastRow;              //2020/10/12 0:00:00
            string[] fileName = Directory.GetFiles(System.Environment.CurrentDirectory + "\\ExportData", "*.csv");

            string firstDate = DateTime.Now.Year.ToString() + "/" + fileName[0].Split("TMAD-")[1].Split(".")[0].Split("-")[0] + "/" + fileName[0].Split("TMAD-")[1].Split(".")[0].Split("-")[1]; //获取文件列表中第一项文件日期
            //string lastDate = DateTime.Now.Year.ToString() + "/" + fileName[^1].Split("Total-")[1].Split(".")[0].Split("-")[0] + "/" + fileName[^1].Split("Total-")[1].Split(".")[0].Split("-")[1];

            string[] date_1 = sheet.Range["B" + lastRow].Value == "Date" ? firstDate.Split("/") : sheet.Range["B" + lastRow].Value.Split(" ")[0].ToString().Split("/");
            DateTime dateTime = new DateTime(Convert.ToInt32(date_1[0]), Convert.ToInt32(date_1[1]), Convert.ToInt32(date_1[2]));

            string curr_Date = DateTime.Now.ToString("MM-dd");
            string curr_Date1 = DateTime.Now.ToString("yyyy/M/dd");


            for (int i = 1; i < fileName.Length + 10; i++)
            {
                int lastRow1 = sheet.Range.LastRow + 1;

                string new_date1 = dateTime.AddDays(i).ToString("MM-dd");
                string fullpath = path + Crashes_Total + new_date1 + suffix;
                string fullpath1 = path + TMAD + new_date1 + suffix;
                DateTime new_date = dateTime.AddDays(i);
                if (fileName.Contains(fullpath))
                {
                    sheet.Range["B" + lastRow1].DateTimeValue = new_date;
                    for (int x = 1; x <= name.Length; x++)
                    {
                        sheet.Range[Convert.ToChar('B' + x).ToString() + lastRow1].NumberValue = AnalyzeCSV(fullpath, "Crashes", name[x - 1], condition)[0];
                    }
                    sheet.Range[Convert.ToChar('B' + (name.Length + 1)).ToString() + lastRow1].NumberValue = AnalyzeCSV(fullpath, "Crashes", name[0], condition)[1];

                    if (fileName.Contains(fullpath1))
                    {
                        if (AnalyzeCSV(fullpath1, "TMAD", "2004 | Vb", condition)[0] != 0)
                        {
                            sheet.Range[Convert.ToChar('B' + (name.Length + 4)).ToString() + lastRow1].NumberValue = AnalyzeCSV(fullpath1, "TMAD", "2004 | Vb", condition)[0];
                        }
                        else
                        {
                            sheet.Range[Convert.ToChar('B' + (name.Length + 4)).ToString() + lastRow1].NumberValue = AnalyzeCSV(fullpath1, "TMAD", "Insider | Vb", condition)[0];
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

                if (new_date.ToString("yyyy/M/dd") == curr_Date1 || dateTime.ToString("yyyy/M/dd") == curr_Date1) { break; }
            }
            Console.WriteLine("now date :" + curr_Date);
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

                workbook1.Worksheets["Crashes"].Range["B6:" + Convert.ToChar('B' + (name.Length + 4)).ToString() + "6"].Copy(workbook.Worksheets["Crashes"].Range["A6:" + Convert.ToChar('A' + (name.Length + 4)).ToString() + "6"]);
                if (lastRow < 37)
                {
                    workbook1.Worksheets["Crashes"].Range["B7" + ":" + Convert.ToChar('B' + (name.Length + 4)).ToString() + lastRow].Copy(workbook.Worksheets["Crashes"].Range["A7:" + Convert.ToChar('A' + (name.Length + 4)).ToString() + "36"]);
                }
                else
                {
                    workbook1.Worksheets["Crashes"].Range["B" + (lastRow - 29) + ":" + Convert.ToChar('B' + (name.Length + 4)).ToString() + lastRow].Copy(workbook.Worksheets["Crashes"].Range["A7:" + Convert.ToChar('A' + (name.Length + 4)).ToString() + "36"]);

                }

                workbook.Worksheets["Crashes"].Range["A5:H36"].HorizontalAlignment = HorizontalAlignType.Center;
                workbook.SaveToFile(System.Environment.CurrentDirectory + "\\Trend Analysis.xlsx", ExcelVersion.Version2013);

            }
            catch (Exception)
            {
                throw;
            }
        }

        public static int[] AnalyzeCSV(string path, string type, string Name, Dictionary<string, string[]> condition)
        {
            try
            {
                bool flag = readCSV(path, out DataTable dt);
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
                    //crashes[1] = dt.AsEnumerable().Where(d => Convert.ToInt32(d.Field<string>("Crashes"))).Sum();
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
                                //Tmad[0] += dt.AsEnumerable().Where(d => d["OSVersion"].ToString() == item).Select(d => Convert.ToInt32(d.Field<string>("[TMAD]"))).Sum();
                            }
                        }
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
        public static List<string> GetItems()
        {

            string[] fileName = Directory.GetFiles(Environment.CurrentDirectory + "\\ExportData", "*.csv");
            readCSV(fileName[^1], out DataTable dt);
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
                //dt.AsEnumerable().Where(d=>d["DriverName"].ToString()== condition)
                List<string> items = dt.AsEnumerable().Where(d => d["DriverName"].ToString() == "rltkapou64.dll").Select(d => d.Field<string>("OSVersion")).Distinct().ToList<string>();
                items.Sort();
                return items;
            }
            catch (Exception)
            {

                throw;
            }
        }
        public static int GetCrashes1(DataTable d, string name, Dictionary<string, string[]> condition)
        {
            List<string> lists = GetItems();
            int crashes = 0;
            StringBuilder exp = new StringBuilder();
            exp.AppendFormat(" DriverName ='{0}'", name);
            foreach (var values in condition)
            {
                if (values.Key.Equals("ReleaseVersion"))
                {
                    int i = 0;
                    exp.Append(" and (");
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
                    exp.Append(" and (");
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
                if (values.Key.Equals("DriverVersion"))
                {
                    int y = 0;
                    exp.Append(" and (");
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
                crashes = dataRows.Select(d => Convert.ToInt32(d.Field<string>("Crashes"))).Sum();
                return crashes;
            }
            catch (Exception)
            {
                throw;
            }
        }
        public static int GetCrashes(DataTable d, string name, Dictionary<string, string[]> condition)
        {
            StringBuilder exp = new StringBuilder();
            DataTable dt = new DataTable();
            DataTable dt1 = new DataTable();

            int crashes = 0;
            EnumerableRowCollection<DataRow> enumerable = dt.AsEnumerable();
            try
            {
                dt = d.AsEnumerable().Where(d => d["DriverName"].ToString() == name).CopyToDataTable();
                if (d.Columns.Count > 0)
                {
                    foreach (DataColumn drVal in d.Columns)
                    {
                        dt1.Columns.Add(drVal.ColumnName);
                    }
                }
                foreach (var values in condition)
                {
                    if (values.Key.Equals("ReleaseVersion"))
                    {
                        DataRow[] dataRows = null;
                        foreach (var item in values.Value)
                        {
                            dataRows = dt.AsEnumerable().Where(d => d["ReleaseVersion"].ToString() == item).ToArray();
                            //int p = dt.AsEnumerable().Where(d => d["ReleaseVersion"].ToString() == item).Select(d => Convert.ToInt32(d.Field<string>("Crashes"))).Sum();
                            //crashes += p;
                            if (dataRows.Length > 0)
                            {
                                //dt1.Clear();
                                foreach (DataRow drVal in dataRows)
                                {
                                    dt1.ImportRow(drVal);
                                }
                            }
                        }
                    }
                    if (values.Key.Equals("OSVersion"))
                    {
                        DataRow[] dataRows1 = null;
                        int i = 0;
                        foreach (var item in values.Value)
                        {
                            dataRows1 = dt1.AsEnumerable().Where(d => d["OSVersion"].ToString() == item).ToArray();
                            if (i == 0)
                            {
                                dt1.Rows.Clear();
                            }
                            i++;
                            foreach (DataRow drVal in dataRows1)
                            {
                                dt1.ImportRow(drVal);
                            }
                        }
                    }
                    if (values.Key.Equals("DriverVersion"))
                    {
                        DataRow[] dataRows2 = null;
                        foreach (var item in values.Value)
                        {
                            dataRows2 = dt1.AsEnumerable().Where(d => d["DriverVersion"].ToString() == item).ToArray();
                            if (dataRows2.Length > 0)
                            {
                                //dt1.Clear();
                                foreach (DataRow drVal in dataRows2)
                                {
                                    dt1.ImportRow(drVal);
                                }
                            }
                        }
                    }
                    int count1 = dt1.Rows.Count;
                    crashes = dt1.AsEnumerable().Select(d => Convert.ToInt32(d.Field<string>("Crashes"))).Sum();
                }
                return crashes;
            }
            catch (Exception)
            {
                throw;
            }
        }

        public static bool readCSV1(string filePath, out DataTable dt)//从csv读取数据返回table
        {
            dt = new DataTable();
            try
            {
                System.Text.Encoding encoding = Encoding.Default;//GetType(filePath); //
                                                                 // DataTable dt = new DataTable();
                System.IO.FileStream fs = new System.IO.FileStream(filePath, System.IO.FileMode.Open,
                    System.IO.FileAccess.Read);


                System.IO.StreamReader sr = new System.IO.StreamReader(fs, encoding);
                //记录每次读取的一行记录
                string strLine = "";
                //记录每行记录中的各字段内容
                string[] aryLine = null;
                string[] tableHead = null;
                int columnCount = 0;
                //标示是否是读取的第一行
                bool IsFirst = true;
                //逐行读取CSV中的数据
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
                if (aryLine != null && aryLine.Length > 0)
                {
                    dt.DefaultView.Sort = tableHead[0] + " " + "asc";
                }
                sr.Close();
                fs.Close();
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}

