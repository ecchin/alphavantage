using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
using ServiceStack;
using ServiceStack.Text;
using System.IO;
using System.Net.Http;
using Microsoft.Win32;
using System.Xml;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using System.Collections;
using System.Reflection;
using System.Data.OleDb;
using System.Xml.Serialization;
using System.Data;
using System.Threading;
using NPOI.SS.UserModel;
using NPOI.HSSF.Model; // InternalWorkbook
using NPOI.HSSF.UserModel; // HSSFWorkbook, HSSFSheet
using NPOI.XSSF.Model; // InternalWorkbook
using NPOI.XSSF.UserModel; // HSSFWorkbook, HSSFSheet


namespace alphavantage
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        string stockl = null;
        string size = null;
        string min = null;
        string period = null;
        string format = null;
        public MainWindow()
        {
            InitializeComponent();
        }

        public void Getstockdata()
        {
            try
            {
                XmlDocument xml = new XmlDocument();
                String exepath = AppDomain.CurrentDomain.BaseDirectory;
                xml.Load(exepath + @"config.xml");

                XmlNode tokenst = xml.SelectSingleNode("/configuration/token");
                string token = tokenst.InnerText;

                StreamReader stocklist = new StreamReader("stocklist.txt", Encoding.Default);

                while ((stockl = stocklist.ReadLine()) != null)
                {

                    textb1.Dispatcher.BeginInvoke((Action)(() => textb1.Text = "Symbol: " + stockl));
                    var symbol = stockl;


                    if (size.Equals("TIME_SERIES_DAILY"))
                    {
                        var dailyprices = $"https://www.alphavantage.co/query?function=TIME_SERIES_DAILY&symbol={symbol}&outputsize={period}&apikey={token}&datatype=csv"
                        .GetStringFromUrl().FromCsv<List<Program.AlphaVantageData>>();

                        if (format.Equals("tab"))
                        {
                            //StreamWriter connect5s = new StreamWriter(@symbol + "_daily_" + DateTime.Now.ToString().Replace("/", "_").Replace(" ", "_").Replace(":", "_") + ".txt", false, Encoding.Default);
                            StreamWriter connect5s = new StreamWriter(@symbol + "_" + size + "_" + min + "_" + DateTime.Now.ToString().Replace("/", "_").Replace(" ", "_").Replace(":", "_") + ".txt", false, Encoding.Default);
                            for (int x = 0; x < dailyprices.Count; x++)
                            {
                                connect5s.WriteLine(symbol + "\t" + dailyprices[x].Timestamp + "\t" + dailyprices[x].Volume + "\t" + dailyprices[x].Open + "\t" + dailyprices[x].High + "\t" + dailyprices[x].Low + "\t" + dailyprices[x].Close);
                            }
                            connect5s.Close();
                        }

                        else if (format.Equals("xls"))
                        {
                            HSSFWorkbook wb;
                            HSSFSheet sh;

                            wb = new HSSFWorkbook();

                            // create sheet
                            sh = (HSSFSheet)wb.CreateSheet("Sheet1");

                            IDataFormat dataFormatCustom = wb.CreateDataFormat();
                            ICellStyle style1 = wb.CreateCellStyle();
                            style1.DataFormat = dataFormatCustom.GetFormat("MM/dd/yyyy HH:mm:ss AM/PM");

                            for (int i = 0; i < dailyprices.Count; i++)
                            {
                                var r = sh.CreateRow(i);
                                for (int j = 0; j < 7; j++)
                                {
                                    IRow row = sh.GetRow(i);
                                    if (j == 0)
                                    {
                                        r.CreateCell(j);
                                        ICell cell1 = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1.SetCellValue(symbol);
                                    }

                                    else if (j == 1)
                                    {
                                        r.CreateCell(j).CellStyle = style1;
                                        ICell cell1 = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1.SetCellValue(dailyprices[i].Timestamp);
                                    }

                                    else if (j == 2)
                                    {
                                        r.CreateCell(j);
                                        ICell cell1 = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1.SetCellValue((double)dailyprices[i].Volume);
                                    }

                                    else if (j == 3)
                                    {
                                        r.CreateCell(j);
                                        ICell cell1 = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1.SetCellValue((double)dailyprices[i].Open);
                                    }

                                    else if (j == 4)
                                    {
                                        r.CreateCell(j);
                                        ICell cell1 = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1.SetCellValue((double)dailyprices[i].High);
                                    }

                                    else if (j == 5)
                                    {
                                        r.CreateCell(j);
                                        ICell cell1 = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1.SetCellValue((double)dailyprices[i].Low);
                                    }

                                    else if (j == 6)
                                    {
                                        r.CreateCell(j);
                                        ICell cell1 = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1.SetCellValue((double)dailyprices[i].Close);
                                    }

                                }
                            }

                            //using (var fs = new FileStream(@symbol + "_daily_" + DateTime.Now.ToString().Replace("/", "_").Replace(" ", "_").Replace(":", "_") + ".xls", FileMode.Create, FileAccess.Write))
                            using (var fs = new FileStream(@symbol + "_" + size + "_" + min + "_" + DateTime.Now.ToString().Replace("/", "_").Replace(" ", "_").Replace(":", "_") + ".xls", FileMode.Create, FileAccess.Write))
                            {
                                wb.Write(fs);
                            }

                        }

                        else if (format.Equals("xlsx"))
                        {
                            XSSFWorkbook wb;
                            XSSFSheet sh;
                            wb = new XSSFWorkbook();

                            // create sheet
                            sh = (XSSFSheet)wb.CreateSheet("Sheet1");
                            IDataFormat dataFormatCustom = wb.CreateDataFormat();
                            ICellStyle style1 = wb.CreateCellStyle();
                            style1.DataFormat = dataFormatCustom.GetFormat("MM/dd/yyyy HH:mm:ss AM/PM");
                            for (int i = 0; i < dailyprices.Count; i++)
                            {
                                var r = sh.CreateRow(i);
                                for (int j = 0; j < 7; j++)
                                {
                                    IRow row = sh.GetRow(i);
                                    if (j == 0)
                                    {
                                        r.CreateCell(j);
                                        ICell cell1 = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1.SetCellValue(symbol);
                                    }

                                    else if (j == 1)
                                    {
                                        r.CreateCell(j).CellStyle = style1;
                                        ICell cell1 = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1.SetCellValue(dailyprices[i].Timestamp);
                                    }

                                    else if (j == 2)
                                    {
                                        r.CreateCell(j);
                                        ICell cell1 = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1.SetCellValue((double)dailyprices[i].Volume);
                                    }

                                    else if (j == 3)
                                    {
                                        r.CreateCell(j);
                                        ICell cell1 = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1.SetCellValue((double)dailyprices[i].Open);
                                    }

                                    else if (j == 4)
                                    {
                                        r.CreateCell(j);
                                        ICell cell1 = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1.SetCellValue((double)dailyprices[i].High);
                                    }

                                    else if (j == 5)
                                    {
                                        r.CreateCell(j);
                                        ICell cell1 = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1.SetCellValue((double)dailyprices[i].Low);
                                    }

                                    else if (j == 6)
                                    {
                                        r.CreateCell(j);
                                        ICell cell1 = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1.SetCellValue((double)dailyprices[i].Close);
                                    }

                                }
                            }

                            
                            using (var fs = new FileStream(@symbol + "_" + size + "_" + min + "_" + DateTime.Now.ToString().Replace("/", "_").Replace(" ", "_").Replace(":", "_") + ".xlsx", FileMode.Create, FileAccess.Write))
                            {
                                wb.Write(fs);
                            }

                        }

                    }

                    else
                    {


                        var dailyprices = $"https://www.alphavantage.co/query?function=TIME_SERIES_INTRADAY&symbol={symbol}&outputsize={period}&interval={min}&apikey={token}&datatype=csv"
                       .GetStringFromUrl().FromCsv<List<Program.AlphaVantageData>>();

                        if (format.Equals("tab"))
                        {
                            //StreamWriter connect5s = new StreamWriter(@symbol + "_intraday_" + min + "_" + DateTime.Now.ToString().Replace("/", "_").Replace(" ", "_").Replace(":", "_") + ".txt", false, Encoding.Default);
                            StreamWriter connect5s = new StreamWriter(@symbol + "_" + size + "_" + min + "_" + DateTime.Now.ToString().Replace("/", "_").Replace(" ", "_").Replace(":", "_") + ".txt", false, Encoding.Default);

                            for (int x = 0; x < dailyprices.Count; x++)
                            {
                                connect5s.WriteLine(symbol + "\t" + dailyprices[x].Timestamp + "\t" + dailyprices[x].Volume + "\t" + dailyprices[x].Open + "\t" + dailyprices[x].High + "\t" + dailyprices[x].Low + "\t" + dailyprices[x].Close);
                            }

                            connect5s.Close();
                        }

                        else if (format.Equals("xls"))
                        {
                            HSSFWorkbook wb;
                            HSSFSheet sh;

                            wb = new HSSFWorkbook();

                            // create sheet
                            sh = (HSSFSheet)wb.CreateSheet("Sheet1");
                            IDataFormat dataFormatCustom = wb.CreateDataFormat();
                            ICellStyle style1 = wb.CreateCellStyle();
                            style1.DataFormat = dataFormatCustom.GetFormat("MM/dd/yyyy HH:mm:ss AM/PM");

                            for (int i = 0; i < dailyprices.Count; i++)
                            {
                                var r = sh.CreateRow(i);
                                for (int j = 0; j < 7; j++)
                                {
                                    IRow row = sh.GetRow(i);
                                    if (j == 0)
                                    {
                                        r.CreateCell(j);
                                        ICell cell1 = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1.SetCellValue(symbol);
                                    }

                                    else if (j == 1)
                                    {
                                        r.CreateCell(j).CellStyle = style1;
                                        ICell cell1 = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1.SetCellValue(dailyprices[i].Timestamp);
                                    }

                                    else if (j == 2)
                                    {
                                        r.CreateCell(j);
                                        ICell cell1 = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1.SetCellValue((double)dailyprices[i].Volume);
                                    }

                                    else if (j == 3)
                                    {
                                        r.CreateCell(j);
                                        ICell cell1 = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1.SetCellValue((double)dailyprices[i].Open);
                                    }

                                    else if (j == 4)
                                    {
                                        r.CreateCell(j);
                                        ICell cell1 = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1.SetCellValue((double)dailyprices[i].High);
                                    }

                                    else if (j == 5)
                                    {
                                        r.CreateCell(j);
                                        ICell cell1 = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1.SetCellValue((double)dailyprices[i].Low);
                                    }

                                    else if (j == 6)
                                    {
                                        r.CreateCell(j);
                                        ICell cell1 = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1.SetCellValue((double)dailyprices[i].Close);
                                    }

                                }
                            }

                            //using (var fs = new FileStream(@symbol + "_intraday_" + min + "_" + DateTime.Now.ToString().Replace("/", "_").Replace(" ", "_").Replace(":", "_") + ".xls", FileMode.Create, FileAccess.Write))
                            using (var fs = new FileStream(@symbol + "_" + size + "_" + min + "_" + DateTime.Now.ToString().Replace("/", "_").Replace(" ", "_").Replace(":", "_") + ".xls", FileMode.Create, FileAccess.Write))
                            {
                                wb.Write(fs);
                            }

                        }

                        else if (format.Equals("xlsx"))
                        {
                            XSSFWorkbook wb;
                            XSSFSheet sh;
                            wb = new XSSFWorkbook();

                            // create sheet
                            sh = (XSSFSheet)wb.CreateSheet("Sheet1");
                            IDataFormat dataFormatCustom = wb.CreateDataFormat();
                            ICellStyle style1 = wb.CreateCellStyle();
                            style1.DataFormat = dataFormatCustom.GetFormat("MM/dd/yyyy HH:mm:ss AM/PM");

                            for (int i = 0; i < dailyprices.Count; i++)
                            {
                                var r = sh.CreateRow(i);
                                for (int j = 0; j < 7; j++)
                                {
                                    IRow row = sh.GetRow(i);
                                    if (j == 0)
                                    {
                                        r.CreateCell(j);
                                        ICell cell1 = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1.SetCellValue(symbol);
                                    }

                                    else if (j == 1)
                                    {
                                        r.CreateCell(j).CellStyle = style1;
                                        ICell cell1 = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1.SetCellValue(dailyprices[i].Timestamp);
                                    }

                                    else if (j == 2)
                                    {
                                        r.CreateCell(j);
                                        ICell cell1 = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1.SetCellValue((double)dailyprices[i].Volume);
                                    }

                                    else if (j == 3)
                                    {
                                        r.CreateCell(j);
                                        ICell cell1 = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1.SetCellValue((double)dailyprices[i].Open);
                                    }

                                    else if (j == 4)
                                    {
                                        r.CreateCell(j);
                                        ICell cell1 = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1.SetCellValue((double)dailyprices[i].High);
                                    }

                                    else if (j == 5)
                                    {
                                        r.CreateCell(j);
                                        ICell cell1 = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1.SetCellValue((double)dailyprices[i].Low);
                                    }

                                    else if (j == 6)
                                    {
                                        r.CreateCell(j);
                                        ICell cell1 = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                        cell1.SetCellValue((double)dailyprices[i].Close);
                                    }

                                }
                            }

                            //using (var fs = new FileStream(@symbol + "_intraday_" + min + "_" + DateTime.Now.ToString().Replace("/", "_").Replace(" ", "_").Replace(":", "_") + ".xlsx", FileMode.Create, FileAccess.Write))
                            using (var fs = new FileStream(@symbol + "_" + size + "_" + min + "_" + DateTime.Now.ToString().Replace("/", "_").Replace(" ", "_").Replace(":", "_") + ".xlsx", FileMode.Create, FileAccess.Write))
                            {
                                wb.Write(fs);
                            }

                        }

                    }

                }

                stocklist.Close();
            }

            catch (Exception e)
            {
                Console.WriteLine($"Generic Exception Handler: {e}");
            }

            finally
            {

            }
        }


        static void Getstockind()
        {
            XmlDocument xml = new XmlDocument();
            String exepath = AppDomain.CurrentDomain.BaseDirectory;
            xml.Load(exepath + @"config.xml");

            XmlNode tokenst = xml.SelectSingleNode("/configuration/token");
            string token = tokenst.InnerText;


            var symbol = "MSFT";
            StreamWriter connect7 = new StreamWriter(@"C:\test\stockapi\alphavantage\bin\Debug\" + symbol + "_60minRSI.txt", false, Encoding.Default);

            var RSI60min = $"https://www.alphavantage.co/query?function=RSI&symbol=MSFT&interval=60min&time_period=20&series_type=close&apikey={token}&datatype=csv"
                    .GetStringFromUrl();

            string[] RSIele = RSI60min.TrimEnd('\n').Split('\n');

            for (int x = 0; x < RSIele.Length; x++)
            {
                Console.WriteLine(RSIele[x]);              
            }
        }

        private async void Download_Click(object sender, RoutedEventArgs e)
        {

            download.IsEnabled = false;

            if (timeframe.SelectionBoxItem.ToString().Equals("Daily"))
            {
                size = "TIME_SERIES_DAILY";
                min = "";
            }

            else if (timeframe.SelectionBoxItem.ToString().Equals("60 min"))
            {
                size = "TIME_SERIES_INTRADAY";
                min = "60min";
            }

            else if (timeframe.SelectionBoxItem.ToString().Equals("30 min"))
            {
                size = "TIME_SERIES_INTRADAY";
                min = "30min";
            }

            else if (timeframe.SelectionBoxItem.ToString().Equals("15 min"))
            {
                size = "TIME_SERIES_INTRADAY";
                min = "15min";
            }

            else if (timeframe.SelectionBoxItem.ToString().Equals("5 min"))
            {
                size = "TIME_SERIES_INTRADAY";
                min = "5min";
            }

            else if (timeframe.SelectionBoxItem.ToString().Equals("1 min"))
            {
                size = "TIME_SERIES_INTRADAY";
                min = "1min";
            }

            else
            {
                MessageBox.Show("invalid time series");
            }
      

            if (datap.SelectionBoxItem.ToString().Equals("Full"))
            {
                period = "full";
            }

            else if (datap.SelectionBoxItem.ToString().Equals("Compact"))
            {
                period = "compact";
            }

            else
            {
                MessageBox.Show("invalid period");
            }
  
            if (filetp.SelectionBoxItem.ToString().Equals("tab delimited"))
            {
                format = "tab";
            }

            else if (filetp.SelectionBoxItem.ToString().Equals("excel xls"))
            {
                format = "xls";
            }

            else if (filetp.SelectionBoxItem.ToString().Equals("excel xlsx"))
            {
                format = "xlsx";
            }

            else
            {
                MessageBox.Show("invalid format");
            }

            await Task.Run(() =>
            {
                Getstockdata();

            });


            textb1.Text = "Download Completed!";
            download.IsEnabled = true;
        }

     
    }
}
