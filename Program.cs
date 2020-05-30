using System;
using System.Windows;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
    class Program
    {
        [DllImport("kernel32.dll", SetLastError = true, ExactSpelling = true)]
        static extern bool FreeConsole();

        public class AlphaVantageData
        {
            public DateTime Timestamp { get; set; }
            public decimal Open { get; set; }

            public decimal High { get; set; }
            public decimal Low { get; set; }

            public decimal Close { get; set; }
            public decimal Volume { get; set; }
        }

        [STAThread]
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                FreeConsole();
                var app = new MainWindow();
                var application = new System.Windows.Application();              
                application.Run(app);
                return;
            }

            else
            {
                if (args[0].Equals("dldata"))
                {
                    Getstockdata();
                }
                
                //Getstockind();
                //complete message

            }
        }
        
        static async void Getstockdataasync()
        {
            XmlDocument xml = new XmlDocument();
            String exepath = AppDomain.CurrentDomain.BaseDirectory;
            xml.Load(exepath + @"config.xml");

            XmlNode tokenst = xml.SelectSingleNode("/configuration/token");
            string token = tokenst.InnerText;

            StreamReader stocklist = new StreamReader("stocklist.txt", Encoding.Default);
            StreamWriter connect5 = new StreamWriter(@"stockdata_daily.txt", true, Encoding.Default);
            StreamWriter connect6 = new StreamWriter(@"summary.txt", true, Encoding.Default);

            string stockl = null;

            while ((stockl = stocklist.ReadLine()) != null)
            {
                var symbol = stockl;
                StreamWriter connect5s = new StreamWriter(@symbol + "_daily.txt", true, Encoding.Default);
                var dailyprices = new List<AlphaVantageData>();

                await Task.Run(() =>
                {

                    dailyprices = $"https://www.alphavantage.co/query?function=TIME_SERIES_DAILY&symbol={symbol}&outputsize=full&apikey={token}&datatype=csv"
                     .GetStringFromUrl().FromCsv<List<AlphaVantageData>>();

                });

                List<string> bm = new List<string>();
                HSSFWorkbook wb;
                HSSFSheet sh;


                if (!File.Exists(@symbol + "_daily.txt" + "test.xls"))
                {
                    wb = HSSFWorkbook.Create(InternalWorkbook.CreateWorkbook());
               
                    // create sheet
                    sh = (HSSFSheet)wb.CreateSheet("Sheet1");
                  
                    IDataFormat dataFormatCustom = wb.CreateDataFormat();
                    ICellStyle style1 = wb.CreateCellStyle();
                    style1.DataFormat = dataFormatCustom.GetFormat("MM/dd/yyyy HH:mm:ss AM/PM");

                    for (int x = 0; x < dailyprices.Count; x++)
                    {
                        var r = sh.CreateRow(x);

                        for (int j = 0; j < 7; j++)
                        {
                            IRow row = sh.GetRow(x);
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
                                cell1.SetCellValue(dailyprices[x].Timestamp);                               
                            }

                            else if (j == 2)
                            {
                                r.CreateCell(j);
                                ICell cell1 = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                cell1.SetCellValue((double)dailyprices[x].Volume);
                            }

                            else if (j == 3)
                            {
                                r.CreateCell(j);
                                ICell cell1 = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                cell1.SetCellValue((double)dailyprices[x].Open);
                            }

                            else if (j == 4)
                            {
                                r.CreateCell(j);
                                ICell cell1 = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                cell1.SetCellValue((double)dailyprices[x].High);
                            }

                            else if (j == 5)
                            {
                                r.CreateCell(j);
                                ICell cell1 = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                cell1.SetCellValue((double)dailyprices[x].Low);
                            }

                            else if (j == 6)
                            {
                                r.CreateCell(j);
                                ICell cell1 = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                cell1.SetCellValue((double)dailyprices[x].Close);
                            }
                        }
                    }

                    using (var fs = new FileStream(@symbol + "_daily.txt" + "test.xls", FileMode.Create, FileAccess.Write))
                    {
                        wb.Write(fs);
                    }

                }

                for (int x = 0; x < dailyprices.Count; x++)
                {
                    connect5.WriteLine(symbol + "\t" + dailyprices[x].Timestamp + "\t" + dailyprices[x].Volume + "\t" + dailyprices[x].Open + "\t" + dailyprices[x].High + "\t" + dailyprices[x].Low + "\t" + dailyprices[x].Close);
                    connect5s.WriteLine(symbol + "\t" + dailyprices[x].Timestamp + "\t" + dailyprices[x].Volume + "\t" + dailyprices[x].Open + "\t" + dailyprices[x].High + "\t" + dailyprices[x].Low + "\t" + dailyprices[x].Close);
                }

                decimal ten = ((dailyprices[0].Close - dailyprices[9].Close) / (dailyprices[9].Close)) * 100;
                decimal thirty = ((dailyprices[0].Close - dailyprices[29].Close) / (dailyprices[29].Close)) * 100;
                decimal sixty = ((dailyprices[0].Close - dailyprices[59].Close) / (dailyprices[59].Close)) * 100;
                decimal ninety = ((dailyprices[0].Close - dailyprices[89].Close) / (dailyprices[89].Close)) * 100;
                decimal N120 = ((dailyprices[0].Close - dailyprices[119].Close) / (dailyprices[119].Close)) * 100;
                decimal N240 = ((dailyprices[0].Close - dailyprices[239].Close) / (dailyprices[239].Close)) * 100;
                decimal N360 = ((dailyprices[0].Close - dailyprices[359].Close) / (dailyprices[359].Close)) * 100;

                connect6.WriteLine(symbol + " 10 day " + ten);
                connect6.WriteLine(symbol + " 30 day " + thirty);
                connect6.WriteLine(symbol + " 60 day " + sixty);
                connect6.WriteLine(symbol + " 90 day " + ninety);
                connect6.WriteLine(symbol + " 120 day " + N120);
                connect6.WriteLine(symbol + " 240 day " + N240);
                connect6.WriteLine(symbol + " 360 day " + N360);

                connect5s.Close();
            }
            
            connect5.Close();
            connect6.Close();
            stocklist.Close();

        }

        static void Getstockdata()
        {
            //main1
            //main2
            //main3



            try
            {
                XmlDocument xml = new XmlDocument();
                String exepath = AppDomain.CurrentDomain.BaseDirectory;
                xml.Load(exepath + @"config.xml");

                XmlNode tokenst = xml.SelectSingleNode("/configuration/token");
                string token = tokenst.InnerText;

                XmlNode functionst = xml.SelectSingleNode("/configuration/function");
                string function = functionst.InnerText;

                XmlNode outputsizest = xml.SelectSingleNode("/configuration/outputsize");
                string outputsize = outputsizest.InnerText;

                XmlNode intervalst = xml.SelectSingleNode("/configuration/interval");
                string interval = intervalst.InnerText;


                StreamReader stocklist = new StreamReader(@AppDomain.CurrentDomain.BaseDirectory + "stocklist.txt", Encoding.Default);
             
                string stockl = null;

                while ((stockl = stocklist.ReadLine()) != null)
                {
                    var symbol = stockl;

                    //StreamWriter connect5s = new StreamWriter(@symbol + "_daily.txt", true, Encoding.Default);
                    StreamWriter connect5s = new StreamWriter(@symbol + "_" + function + "_" + interval + "_" + DateTime.Now.ToString().Replace("/", "_").Replace(" ", "_").Replace(":", "_") + ".txt", false, Encoding.Default);

                    //var dailyprices = $"https://www.alphavantage.co/query?function=TIME_SERIES_DAILY&symbol={symbol}&outputsize=full&apikey={token}&datatype=csv"
                    // .GetStringFromUrl().FromCsv<List<AlphaVantageData>>();


                    var prices = $"https://www.alphavantage.co/query?function={function}&symbol={symbol}&outputsize={outputsize}&interval={interval}&apikey={token}&datatype=csv"
                        .GetStringFromUrl().FromCsv<List<Program.AlphaVantageData>>();


                    HSSFWorkbook wb;
                    HSSFSheet sh;          
                    wb = HSSFWorkbook.Create(InternalWorkbook.CreateWorkbook());

                    // create sheet
                    sh = (HSSFSheet)wb.CreateSheet("Sheet1");
                    IDataFormat dataFormatCustom = wb.CreateDataFormat();
                    ICellStyle style1 = wb.CreateCellStyle();
                    style1.DataFormat = dataFormatCustom.GetFormat("MM/dd/yyyy HH:mm:ss AM/PM");

                    XSSFWorkbook wbx;
                    XSSFSheet shx;


                    for (int x = 0; x < prices.Count; x++)
                    {
                        var r = sh.CreateRow(x);

                        for (int j = 0; j < 7; j++)
                        {
                            IRow row = sh.GetRow(x);
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
                                cell1.SetCellValue(prices[x].Timestamp);
                            }

                            else if (j == 2)
                            {
                                r.CreateCell(j);
                                ICell cell1 = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                cell1.SetCellValue((double)prices[x].Volume);
                            }

                            else if (j == 3)
                            {
                                r.CreateCell(j);
                                ICell cell1 = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                cell1.SetCellValue((double)prices[x].Open);
                            }

                            else if (j == 4)
                            {
                                r.CreateCell(j);
                                ICell cell1 = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                cell1.SetCellValue((double)prices[x].High);
                            }

                            else if (j == 5)
                            {
                                r.CreateCell(j);
                                ICell cell1 = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                cell1.SetCellValue((double)prices[x].Low);
                            }

                            else if (j == 6)
                            {
                                r.CreateCell(j);
                                ICell cell1 = row.GetCell(j, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                                cell1.SetCellValue((double)prices[x].Close);
                            }
                        }
                    }

                    using (var fs = new FileStream(@symbol + "_" + function + "_" + interval + "_" + DateTime.Now.ToString().Replace("/", "_").Replace(" ", "_").Replace(":", "_") + ".xls", FileMode.Create, FileAccess.Write))
                    {
                        wb.Write(fs);
                    }

                    //}

                    for (int x = 0; x < prices.Count; x++)
                    {
                        //connect5.WriteLine(symbol + "\t" + prices[x].Timestamp + "\t" + prices[x].Volume + "\t" + prices[x].Open + "\t" + prices[x].High + "\t" + prices[x].Low + "\t" + prices[x].Close);
                        connect5s.WriteLine(symbol + "\t" + prices[x].Timestamp + "\t" + prices[x].Volume + "\t" + prices[x].Open + "\t" + prices[x].High + "\t" + prices[x].Low + "\t" + prices[x].Close);
                    }

                    decimal ten = ((prices[0].Close - prices[9].Close) / (prices[9].Close)) * 100;
                    decimal thirty = ((prices[0].Close - prices[29].Close) / (prices[29].Close)) * 100;
                    decimal sixty = ((prices[0].Close - prices[59].Close) / (prices[59].Close)) * 100;
                    decimal ninety = ((prices[0].Close - prices[89].Close) / (prices[89].Close)) * 100;
                    decimal N120 = ((prices[0].Close - prices[119].Close) / (prices[119].Close)) * 100;
                    decimal N240 = ((prices[0].Close - prices[239].Close) / (prices[239].Close)) * 100;
                    decimal N360 = ((prices[0].Close - prices[359].Close) / (prices[359].Close)) * 100;

                    //connect6.WriteLine(symbol + " 10 day " + ten);
                    //connect6.WriteLine(symbol + " 30 day " + thirty);
                    //connect6.WriteLine(symbol + " 60 day " + sixty);
                    //connect6.WriteLine(symbol + " 90 day " + ninety);
                    //connect6.WriteLine(symbol + " 120 day " + N120);
                    //connect6.WriteLine(symbol + " 240 day " + N240);
                    //connect6.WriteLine(symbol + " 360 day " + N360);

                    connect5s.Close();
                }

                //connect5.Close();
                //connect6.Close();
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

    }
}
