using System;
using System.Windows;
using HtmlAgilityPack;
using System.IO;
using EquityDailyWPF.Helper;
using System.Linq;
using System.Globalization;
using EquityDailyWPF;
using System.Diagnostics;
using System.Net;
using System.Net.Mail;
using Microsoft.Office.Interop.Excel;
using static System.Net.WebRequestMethods;
//using static System.Net.WebRequestMethods;

namespace EquityDailyWPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {

        public string niftyDailyUpdateFile;
        public string chinaDailyUpdateFile;
        public string equityDailyUpdateFile;
        public string openlowhighDailyUpdateFile;
        public string CopyToDisplayInGrid;
        public string niftyDailyUpdateOutputDirectoryFile;
        public string chinaDailyUpdateFileOutputDirectory;
        public string equityDailyUpdateFileOutputDirectory;
        public string openlowstrategyyDailyUpdateFileOutputDirectory;

        public HtmlWeb web;
        public string[] stockArrayList;
        int lastUsedRow;
        public string[] stockHighLowURLArrayList;
        public static string[] stockRSIArrayList;

        public string[] stockRSILArrayList;
        public int i = 0;


        //string todayHighOrLowStatusPositiveOrNegativeNasdaq;
        Microsoft.Office.Interop.Excel.Application excel;
        Microsoft.Office.Interop.Excel.Workbook worKbooK;
        Microsoft.Office.Interop.Excel.Worksheet worKsheeT;



        public MainWindow()
        {
            InitializeComponent();
            web = new HtmlWeb();
           // System.Windows.MessageBox.Show("Hi");
            Trace.Listeners.Add(new TextWriterTraceListener("yourlog.log"));
            Trace.AutoFlush = true;
            Trace.Indent();
            Trace.WriteLine("Entering Main..Log started genrating at" + DateTime.Now.ToString());
            Console.WriteLine("Hello World.");



            excel = new Microsoft.Office.Interop.Excel.Application();
            excel.DisplayAlerts = false;


            System.Windows.MessageBox.Show("Hi");

            DirectoryInfo dInfo = Directory.GetParent(Environment.CurrentDirectory);
            dInfo = Directory.GetParent(dInfo.FullName);

            //string dInfo = "D:\\";

            //niftyDailyUpdateFile = dInfo + "\\\\YPH1010144LT\\MyStocks\\DailyniftyView.xlsx";
            niftyDailyUpdateFile = "\\\\YPH1010144LT\\MyStocks\\DailyniftyView.xlsx";

            CopyToDisplayInGrid = "\\\\YPH1010144LT\\DisplayDataInGrid\\Equity.xlsx";

            chinaDailyUpdateFile = dInfo + "\\MyStocks\\china.xlsx";
            //equityDailyUpdateFile = dInfo + "\\MyStocks\\Equity.xlsx";

            equityDailyUpdateFile = "C:\\Users\\butolea\\Downloads\\OptionLovers-master\\EtyPart2.xlsx";
            openlowhighDailyUpdateFile = dInfo + "\\MyStocks\\OpenLowHighStrategy.xlsx";


            niftyDailyUpdateOutputDirectoryFile = dInfo + "\\OutputDaily\\DailyniftyView.xlsx";
            chinaDailyUpdateFileOutputDirectory = dInfo + "\\OutputDaily\\china.xlsx";
           // equityDailyUpdateFileOutputDirectory = dInfo + "\\OutputDaily\\Equity.xlsx";

            equityDailyUpdateFileOutputDirectory = "C:\\Users\\butolea\\Downloads\\OptionLovers-master\\EquityPart2.xlsx";
            openlowstrategyyDailyUpdateFileOutputDirectory = dInfo + "\\OutputDaily\\OpenLowHighStrategy.xlsx";

            //Stocks URl
            stockArrayList = new string[] {  "HPC","PCBL","NT","PhoenixMills",
             "HINDZINC","SRF","SUMICHEM","CROMPTON","CIEINDIA","FINEORG","ALKYLAMINE","ORIENTELEC",
             "INDUSINBANK","NESTLEIND","DIVISLAB","BANDHANBANK","ICICPRLIMITED","PIDILITINDUSTRY","SBICARD","AUROPHARMA","DMART","SWSOLAR",
             "MAHLOG","TVTODAY","DELTACORP","DCBBANK","DHANI","Zomato","IndianHotel","ZenTech","Indigo","EasyTripPlanner"

            };

            //stockArrayList = new string[] { "FCONSUMER", "sbin", "PNB" , "YESBANK"
            //};
            Debug.WriteLine("stockArrayList count is" + stockArrayList.Length);
            //stockRSIArrayList = new string[] {
            //    @"https://trendlyne.com/equity/technical-analysis/FCONSUMER/408/future-consumer-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/SBIN/1193/state-bank-of-india/",
            //    @"https://trendlyne.com/equity/technical-analysis/PNB/1048/punjab-national-bank/",
            //    @"https://trendlyne.com/equity/technical-analysis/YESBANK/1535/yes-bank-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/HDFCBANK/533/hdfc-bank-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/INFY/630/infosys-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/WIPRO/1526/wipro-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/TCS/1372/tata-consultancy-services-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/ZENSARTECH/1544/zensar-technologies-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/SUNPHARMA/1316/sun-pharmaceutical-industries-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/BAJAJ-AUTO/144/bajaj-auto-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/EASUNREYRL/354/easun-reyrolle-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/HCC/528/hindustan-construction-company-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/CIPLA/268/cipla-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/TITAN/1403/titan-company-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/BATAINDIA/168/bata-india-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/RELIANCE/1127/reliance-industries-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/ZEEL/1537/zee-entertainment-enterprises-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/MARUTI/842/maruti-suzuki-india-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/HDFC/532/housing-development-finance-corporation-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/HINDUNILVR/560/hindustan-unilever-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/BPCL/215/bharat-petroleum-corporation-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/ASHOKLEY/114/ashok-leyland-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/AXISBANK/140/axis-bank-ltd/",

            //    @"https://trendlyne.com/equity/technical-analysis/BERGEPAINT/178/berger-paints-india-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/BAJFINANCE/150/bajaj-finance-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/IDBI/588/idbi-bank-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/BHARTIARTL/187/bharti-airtel-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/BOSCHLTD/214/bosch-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/COALINDIA/275/coal-india-ltd/",
               
              


            //    @"https://trendlyne.com/equity/technical-analysis/DABUR/303/dabur-india-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/BIOCON/197/biocon-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/HINDZINC/561/hindustan-zinc-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/ICICIBANK/584/icici-bank-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/IBULHSGFIN/582/indiabulls-housing-finance-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/IOC/639/indian-oil-corporation-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/LT/800/larsen-toubro-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/LICHSGFIN/790/lic-housing-finance-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/NMDC/949/nmdc-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/LUPIN/804/lupin-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/ADVENZYMES/4635/advanced-enzyme-technologies-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/JUBLFOOD/701/jubilant-foodworks-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/CASTROLIND/241/castrol-india-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/TECHM/1374/tech-mahindra-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/TATAELXSI/1358/tata-elxsi-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/CADILAHC/229/cadila-healthcare-ltd/",
            //    @"https://trendlyne.com/equity/1887/NIFTY50/nifty-50/",
            //    @"https://trendlyne.com/equity/technical-analysis/DEN/322/den-networks-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/RBLBANK/4685/rbl-bank-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/BANKBARODA/162/bank-of-baroda/",
            //    @"https://trendlyne.com/equity/technical-analysis/LEMONTREE/81513/lemon-tree-hotels-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/HIMATSEIDE/548/himatsingka-seide-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/KEI/729/kei-industries-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/KEC/727/kec-international-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/NHPC/938/nhpc-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/NTPC/959/ntpc-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/NCC/918/ncc-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/BAJAJFINSV/147/bajaj-finserv-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/TATAMOTORS/1362/tata-motors-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/ITC/647/itc-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/KOTAKBANK/758/kotak-mahindra-bank-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/ONGC/974/oil-and-natural-gas-corporation-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/HCLTECH/531/hcl-technologies-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/DLF/337/dlf-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/AMBUJACEM/71/ambuja-cements-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/IRCTC/167028/indian-railway-catering-tourism-corporation-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/TRIDENT/1415/trident-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/NETWORK18/932/network-18-media-investments-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/FEDERALBNK/412/federal-bank-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/GUJGASLTD/515/gujarat-gas-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/TIPSINDLTD/1401/tips-industries-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/JUMPNET/3552/jump-networks-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/TEJASNET/54898/tejas-networks-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/GULFPETRO/1164/gp-petroleums-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/ASIANPAINT/117/asian-paints-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/UPL/1455/upl-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/CASTROLIND/241/castrol-india-ltd/",
            //    @"https://trendlyne.com/equity/technical-analysis/COLPAL/276/colgate-palmolive-india-ltd/",
            //    @"https://trendlyne.com/equity/1898/NIFTYBANK/nifty-bank/"
            //    //@"",
            //    //@"",
            //    //@"",
            //    //@"",
            //    //@"",
            //    //@"",
            //    //@"",
            //    //@"",
            //    //@"",
            //    //@"",
            //    //@"",
            //    //@"",
            //    //@"",
            //};

           // Debug.WriteLine("stockRSIArrayList count is" + stockRSIArrayList.Length);

            stockHighLowURLArrayList = new string[] {
                @"https://www.moneycontrol.com/india/stockpricequote/refineries/hindustanpetroleumcorporation/HPC",
                @"https://www.moneycontrol.com/india/stockpricequote/carbon-black/pcbl/PCB01",
                @"https://www.moneycontrol.com/india/stockpricequote/software/nazaratechnologies/NT",
                @"https://www.moneycontrol.com/india/stockpricequote/constructioncontracting-real-estate/phoenixmills/PM02",
                @"https://www.moneycontrol.com/india/stockpricequote/metals-non-ferrous/hindustanzinc/HZ",
                @"https://www.moneycontrol.com/india/stockpricequote/diversified/srf/SRF",
                @"https://www.moneycontrol.com/india/stockpricequote/pesticidesagrochemicals/sumitomochemicalindialimite/SUMIC54292",
                @"https://www.moneycontrol.com/india/stockpricequote/electricals/cromptongreavesconsumerelectrical/CGC01",
                @"https://www.moneycontrol.com/india/stockpricequote/auto-ancillaries-castingsforgings/cieautomotiveindia/MF19",
                @"https://www.moneycontrol.com/india/stockpricequote/miscellaneous/fineorganicsindustries/FOI",
                @"https://www.moneycontrol.com/india/stockpricequote/chemicals/alkylamineschemicals/AAC",
                @"https://www.moneycontrol.com/india/stockpricequote/consumer-goods-electronic/orientelectriclimited/ORIEN54130",
                @"https://www.moneycontrol.com/india/stockpricequote/banks-private-sector/indusindbank/IIB",
                @"https://www.moneycontrol.com/india/stockpricequote/food-processing/nestleindia/NI",
                @"https://www.moneycontrol.com/india/stockpricequote/pharmaceuticals/divislaboratories/DL03",
                @"https://www.moneycontrol.com/india/stockpricequote/banks-private-sector/bandhanbank/BB09",
                @"https://www.moneycontrol.com/india/stockpricequote/finance-general/iciciprudentiallifeinsurancecompany/IPL01",
                @"https://www.moneycontrol.com/india/stockpricequote/chemicals/pidiliteindustries/PI11",
                @"https://www.moneycontrol.com/india/stockpricequote/finance-term-lending/sbicardspaymentservices/SCP02",
                @"https://www.moneycontrol.com/india/stockpricequote/pharmaceuticals/aurobindopharma/AP",
                @"https://www.moneycontrol.com/india/stockpricequote/retail/avenuesupermarts/AS19",
                @"https://www.moneycontrol.com/india/stockpricequote/power-generationdistribution/sterlingwilsonrenewableenergy/SWS",
                @"https://www.moneycontrol.com/india/stockpricequote/diversified/mahindralogistics/ML04",
                @"https://www.moneycontrol.com/india/stockpricequote/mediaentertainment/tvtodaynetwork/TVT",
                @"https://www.moneycontrol.com/india/stockpricequote/constructioncontracting-real-estate/deltacorp/DC11",
                @"https://www.moneycontrol.com/india/stockpricequote/banks-private-sector/dcbbank/DCB01",
                @"https://www.moneycontrol.com/india/stockpricequote/finance-investments/dhaniservices/IS13",
                @"https://www.moneycontrol.com/india/stockpricequote/online-services/zomato/Z",
                @"https://www.moneycontrol.com/india/stockpricequote/hotels/indianhotelscompany/IHC",
                @"https://www.moneycontrol.com/india/stockpricequote/defence/zentechnologies/ZT01",
                @"https://www.moneycontrol.com/india/stockpricequote/transportlogistics/interglobeaviation/IA04",
                @"https://www.moneycontrol.com/india/stockpricequote/travel-services/easytripplanners/ETP01",
            };

            Debug.WriteLine("stockHighLowURLArrayList count is" + stockHighLowURLArrayList.Length);


            if (!System.IO.File.Exists(equityDailyUpdateFile))
            {

             //   Microsoft.Office.Interop.Excel.Workbook workbook = new Microsoft.Office.Interop.Excel.Workbook();
                worKbooK = excel.Workbooks.Add(Type.Missing);
                worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;
                worKsheeT.Name = "MyEquity";
                worKsheeT.Cells[1, 1] = "Date";
                worKsheeT.Cells[1, 2] = "Day";
                worKsheeT.Rows[1].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbDarkSalmon;
                int j = 3;
                foreach (var item in stockArrayList)
                {
                    worKsheeT.Cells[1, j] = item.ToString();
                    worKsheeT.Cells[1, ++j] = item.ToString() + "_Status";
                    worKsheeT.Cells[1, ++j] = "RSI";
                    //worKsheeT.Cells[1, ++j] = "MomentumScore" + item.ToString();
                    ++j;
                }
                worKsheeT.Cells.Font.Size = 10;
                worKbooK.SaveAs(equityDailyUpdateFile);
               
                //worKsheeT.Activate();
                //excel.GetSaveAsFilename(equityDailyUpdateFile);
                // worKbooK.Close();
                //  excel.Quit();
                // if (excel != null) { excel.Dispose(); };


            }
          buttonEquityDaily_Click1();
        }

        private async void ButtonEquityDaily_Click()
        { 
        
        }
        private async void buttonEquityDaily_Click1()
        {

            string stockTodaysClosePrice;
            string stockTodayHighOrLowStatus;
            string stockRSIValue;
            bool firstLine = true;
            worKbooK = excel.Workbooks.Open(equityDailyUpdateFile);
            worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;

            Microsoft.Office.Interop.Excel.Range last = worKsheeT.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int lastUsedRow = last.Row;

            var dateAndTime = DateTime.Now;
            var date = dateAndTime.Date;

            worKsheeT.Cells[lastUsedRow + 1, 1] = date;
            worKsheeT.Cells[lastUsedRow + 1, 2] = date.DayOfWeek.ToString();
            // worKsheeT.Cells[lastUsedRow + 1, 2] = "Thursday";

            //This is where stocks actual starts
            int i = 3;
           
            foreach (var stockUrl in stockHighLowURLArrayList.Select((value, zm) => new { zm, value }))
            {
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                var stockUrlvalue = stockUrl.value;
                var index = stockUrl.zm;
                Out_Ref_Params outRefParams = new Out_Ref_Params();
                HtmlDocument loadfirsthtml = web.Load(stockUrlvalue);
                loadfirsthtml.LoadHtml(loadfirsthtml.DocumentNode.InnerHtml.ToString());
                if (loadfirsthtml.DocumentNode.InnerHtml.ToString() != string.Empty)
                    await EquityHelperUtility.StockTodayClosedValueAndStatus(loadfirsthtml, outRefParams, index);

                worKsheeT.Cells[lastUsedRow +1, i] = outRefParams.stockTodayClosedPrice;
                if (DateTime.Now.DayOfWeek.ToString() == "Sunday1")
                {
                    worKsheeT.Cells[lastUsedRow + 2, 2] = "High";
                    worKsheeT.Cells[lastUsedRow + 3, 2] = "Low";
                    worKsheeT.Rows[lastUsedRow + 2].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                    worKsheeT.Rows[lastUsedRow + 3].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                }

                i++;
                worKsheeT.Cells[lastUsedRow + 1, i++] = outRefParams.stockTodayStatus;

                //worKsheeT.Cells[lastUsedRow + 1, i++] = outRefParams.stockRSIValue;

                //if (outRefParams.MomentumScore > 70)
                //{
                //    worKsheeT.Cells[lastUsedRow + 1, i].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbDarkGreen;
                //    worKsheeT.Cells[lastUsedRow + 1, i] = outRefParams.MomentumScore;
                //}
                //else {
                //    if (outRefParams.MomentumScore <= 35)
                //    {
                //        worKsheeT.Cells[lastUsedRow + 1, i].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbOrangeRed;
                //        worKsheeT.Cells[lastUsedRow + 1, i] = outRefParams.MomentumScore;

                //    }
                //    else {
                //        worKsheeT.Cells[lastUsedRow + 1, i] = outRefParams.MomentumScore;

                //    }

                //}

                if (DateTime.Now.DayOfWeek.ToString() == "Sunday1")
                {
                    double[] weekPrice = new double[5];
                    if (!string.IsNullOrEmpty(outRefParams.stockTodayClosedPrice))
                    {
                        weekPrice[0] = float.Parse(outRefParams.stockTodayClosedPrice);
                        int k = 1;
                        double max, min;
                        for (int weekdays = 1; weekdays < 5; weekdays++)
                        {
                            //Trace.WriteLine("I is", + i.ToString());
                            var cellValue = (string)(worKsheeT.Cells[lastUsedRow + 1 - weekdays, i - 3] as Microsoft.Office.Interop.Excel.Range).Value;
                            var cellValue1 = (double)(worKsheeT.Cells[lastUsedRow + 1 - weekdays, i - 2] as Microsoft.Office.Interop.Excel.Range).Value;
                            var cellValue2 = (string)(worKsheeT.Cells[lastUsedRow + 1 - weekdays, i - 1] as Microsoft.Office.Interop.Excel.Range).Value;

                            //weekPrice[k] = float.Parse(worKsheeT.Cells[lastUsedRow + 1 - (weekdays), i - 1]);
                            if ((worKsheeT.Cells[lastUsedRow + 1 - (weekdays), i - 2] as Microsoft.Office.Interop.Excel.Range).Value != null)
                                weekPrice[k] = (double)((worKsheeT.Cells[lastUsedRow + 1 - (weekdays), i - 2] as Microsoft.Office.Interop.Excel.Range).Value);
                            k++;
                        }

                        #region "max and min"
                        max = weekPrice[0];
                        min = weekPrice[0];

                        for (int z = 0; z < 5; z++)
                        {
                            if (weekPrice[z] > max)
                            {
                                max = weekPrice[z];
                                max = weekPrice[z];
                            }


                            if (weekPrice[z] < min)
                            {
                                min = weekPrice[z];
                            }
                        }
                        #endregion
                        worKsheeT.Cells[lastUsedRow + 2, i - 2] = max;
                        worKsheeT.Cells[lastUsedRow + 3, i - 2] = min;

                        string path = @"C:\\Users\\butolea\\Downloads\\OptionLovers-master\\test1.txt";
                        if (!System.IO.File.Exists(path))
                        {
                           var myfile= System.IO.File.Create(path);
                            myfile.Close();
                            TextWriter tw = new StreamWriter(path);
                            tw.WriteLine("The very first line!");
                            tw.Close();
                        }
                        else if (System.IO.File.Exists(path))
                        {
                            using (var tw = new StreamWriter(path, true))
                            {
                                if (firstLine)
                                {
                                    tw.WriteLine("=============================================================================================================================================================================================================================");
                                    tw.WriteLine("Levels By SafeTrader. This is just for Educational purpose.Consult your financial advisor before taking any decision.This Will be great tool for swing trading?brakout trading and adjust portolio. Recommended for Safe Trading Community students only. Stock levels on {0}", DateTime.Today);
                                    firstLine = false;
                                }
                                tw.WriteLine("{0} buy above {1}..sell below {2}", outRefParams.stockName.ToUpper(), max, min);
                                tw.Close();
                            }
                        }
                        worKsheeT.Cells[lastUsedRow + 2, i].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                        worKsheeT.Cells[lastUsedRow + 3, i].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                    }
                }
                i++;

            }

            try
            {
                Trace.WriteLine("Exiting Finally");
                Trace.Unindent();
                Trace.Flush();
                //worKbooK.Close();
               // excel.Quit();
                worKbooK.SaveAs(equityDailyUpdateFile);

                //MailMessage mail = new MailMessage();
                //SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");
                //mail.From = new MailAddress("ankushbutole3@gmail.com");
                //mail.To.Add("optiondiner@gmail.com");
                //mail.Subject = "Test Mail - 1";
                //mail.Body = "mail with attachment";

                //System.Net.Mail.Attachment attachment;
                //attachment = new System.Net.Mail.Attachment("D:\\MyStocks\\test.txt");
                //mail.Attachments.Add(attachment);

                //SmtpServer.Port = 587;
                //SmtpServer.Credentials = new System.Net.NetworkCredential("username", "password");
                //SmtpServer.EnableSsl = true;

                //SmtpServer.Send(mail);
                //MessageBox.Show("mail Send");
            }
            catch (Exception ex)
            {
                //System.Environment.Exit(1);
            }

            //finally {
            //    //Trace.WriteLine("Exiting Finally");
            //    //Trace.Unindent();
            //    //Trace.Flush();

            //}
            System.IO.File.Copy(equityDailyUpdateFile, equityDailyUpdateFileOutputDirectory, true);
            //System.IO.File.Copy(equityDailyUpdateFile, CopyToDisplayInGrid, true);
            worKbooK.Save();
            worKbooK.Close();
            excel.Quit();
            System.Environment.Exit(1);

        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                worKbooK = excel.Workbooks.Open(openlowhighDailyUpdateFile);
                worKsheeT = (Microsoft.Office.Interop.Excel.Worksheet)worKbooK.ActiveSheet;

                Microsoft.Office.Interop.Excel.Range last = worKsheeT.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                lastUsedRow = last.Row;


                foreach (string stockUrl in stockHighLowURLArrayList)
                {
                    HtmlDocument loadfirsthtml = web.Load(stockUrl);
                    loadfirsthtml.LoadHtml(loadfirsthtml.DocumentNode.InnerHtml.ToString());
                    if (loadfirsthtml.DocumentNode.InnerHtml.ToString() != string.Empty)
                        GetTodayHighLow(loadfirsthtml);
                }
                worKbooK.SaveAs(openlowhighDailyUpdateFile);
                System.IO.File.Copy(openlowhighDailyUpdateFile, openlowstrategyyDailyUpdateFileOutputDirectory, true);
                // worKbooK.Save();
                worKbooK.Close();
                excel.Quit();
            }
            catch (Exception ex)
            {
                worKbooK.SaveAs(openlowhighDailyUpdateFile);
                System.IO.File.Copy(openlowhighDailyUpdateFile, openlowstrategyyDailyUpdateFileOutputDirectory, true);
                worKbooK.Close();
                excel.Quit();

            }
        }

        private void GetTodayHighLow(HtmlDocument loadfirsthtml)
        {
            try
            {

                string todayLow = string.Empty;
                string todayHigh = string.Empty;
                string todayOpen = string.Empty;
                string todayVolume = string.Empty;
                float todayCurrentPrice;
                string todayAverageVolume = string.Empty;
                float ThirtyDMA = 0;
                float FiftyDMA = 0;
                float OneFiftyDMA = 0;
                float TwoHundreadthDMA = 0;
                string companyName;
                string previousClose;

                if (i > 1)
                {
                    HtmlNode[] companyNameArray = loadfirsthtml.DocumentNode.SelectNodes("//h1[@class='b_20']").ToArray();
                    companyName = companyNameArray[0].InnerText;
                    HtmlNode[] array = loadfirsthtml.DocumentNode.SelectNodes("//span[@class='gL_11_5']").ToArray();
                    todayLow = array[3].ParentNode.InnerText;
                    todayLow = todayLow.Replace("LOWS: ", string.Empty);
                    //HtmlNode[] todayHigh_Array = loadfirsthtml.DocumentNode.SelectNodes("//span[@id='n_high_sh']").ToArray();
                    todayHigh = array[1].ParentNode.InnerText;
                    todayHigh = todayHigh.Replace("HIGH: ", string.Empty);
                    //HtmlNode[] todayOpen_Array = loadfirsthtml.DocumentNode.SelectNodes("//td[@id='bggry02 br01']").ToArray();
                    todayOpen = array[0].ParentNode.InnerText;
                    todayOpen = todayOpen.Replace("OPEN: ", string.Empty);

                    previousClose = array[2].ParentNode.InnerText;
                    previousClose = previousClose.Replace("PREV CLOSE: ", string.Empty);
                    //HtmlNode[] todayVolume_Array = loadfirsthtml.DocumentNode.SelectNodes("//span[@id='nse_volume']").ToArray();
                    todayVolume = string.Empty;
                    //if (companyName != "NIFTY BANK")
                    //{
                    //When market open
                    // HtmlNode[] todayCurrentPrice_Array = loadfirsthtml.DocumentNode.SelectNodes("//div[@class='FL r_35']").ToArray();

                    //When market is not running
                    HtmlNode[] todayCurrentPrice_Array = loadfirsthtml.DocumentNode.SelectNodes("//div[@class='FL gr_35']").ToArray();
                    todayCurrentPrice = float.Parse(todayCurrentPrice_Array[0].InnerText, CultureInfo.InvariantCulture.NumberFormat);

                    // }
                    //else {
                    //    HtmlNode[] todayCurrentPrice_Array = loadfirsthtml.DocumentNode.SelectNodes("//div[@class='FL gr_35']").ToArray();
                    //    todayCurrentPrice = float.Parse(todayCurrentPrice_Array[0].InnerText, CultureInfo.InvariantCulture.NumberFormat);

                    //}

                    //HtmlNode[] todayAverageVolume_Array = loadfirsthtml.DocumentNode.SelectNodes("//td[@id='avgvol5daysN']").ToArray();
                    todayAverageVolume = string.Empty;
                    HtmlNode[] DMA_Array = loadfirsthtml.DocumentNode.SelectNodes("//td[@class='bb0 br0']").ToArray();
                    ThirtyDMA = float.Parse(DMA_Array[5].InnerText, CultureInfo.InvariantCulture.NumberFormat);
                    FiftyDMA = float.Parse(DMA_Array[7].InnerText, CultureInfo.InvariantCulture.NumberFormat);
                    OneFiftyDMA = float.Parse(DMA_Array[9].InnerText, CultureInfo.InvariantCulture.NumberFormat);
                    TwoHundreadthDMA = float.Parse(DMA_Array[11].InnerText, CultureInfo.InvariantCulture.NumberFormat);
                    // i++;
                }
                else
                {
                    // i++;
                    HtmlNode[] companyNameArray = loadfirsthtml.DocumentNode.SelectNodes("//h1[@class='b_42 company_name']").ToArray();
                    companyName = companyNameArray[0].InnerText;
                    HtmlNode[] todayLow_Array = loadfirsthtml.DocumentNode.SelectNodes("//span[@id='n_low_sh']").ToArray();
                    todayLow = todayLow_Array[0].InnerText;
                    HtmlNode[] todayHigh_Array = loadfirsthtml.DocumentNode.SelectNodes("//span[@id='n_high_sh']").ToArray();
                    todayHigh = todayHigh_Array[0].InnerText;
                    HtmlNode[] todayOpen_Array = loadfirsthtml.DocumentNode.SelectNodes("//div[@id='n_open']").ToArray();
                    todayOpen = todayOpen_Array[0].InnerText;
                    HtmlNode[] todayVolume_Array = loadfirsthtml.DocumentNode.SelectNodes("//span[@id='nse_volume']").ToArray();
                    todayVolume = todayVolume_Array[0].InnerText;
                    HtmlNode[] todayCurrentPrice_Array = loadfirsthtml.DocumentNode.SelectNodes("//span[@id='Nse_Prc_tick']").ToArray();
                    //todayCurrentPrice = todayCurrentPrice_Array[0].InnerText;
                    todayCurrentPrice = float.Parse(todayCurrentPrice_Array[0].InnerText, CultureInfo.InvariantCulture.NumberFormat);

                    HtmlNode[] previousClose_Array = loadfirsthtml.DocumentNode.SelectNodes("//div[@id='n_prevclose']").ToArray();
                    previousClose = previousClose_Array[0].InnerText;

                    HtmlNode[] todayAverageVolume_Array = loadfirsthtml.DocumentNode.SelectNodes("//td[@id='avgvol5daysN']").ToArray();
                    todayAverageVolume = todayAverageVolume_Array[0].InnerText;
                    HtmlNode[] DMA_Array = loadfirsthtml.DocumentNode.SelectNodes("//td[@class='th05 gD_12']").ToArray();
                    //ThirtyDMA =  (DMA_Array[0].InnerText);
                    //FiftyDMA = DMA_Array[1].InnerText;
                    //OneFiftyDMA = DMA_Array[2].InnerText;
                    //TwoHundreadthDMA = DMA_Array[3].InnerText;

                    //old code 
                    //ThirtyDMA = float.Parse(DMA_Array[0].InnerText, CultureInfo.InvariantCulture.NumberFormat);
                    //FiftyDMA = float.Parse(DMA_Array[1].InnerText, CultureInfo.InvariantCulture.NumberFormat);
                    //OneFiftyDMA = float.Parse(DMA_Array[2].InnerText, CultureInfo.InvariantCulture.NumberFormat);
                    //TwoHundreadthDMA = float.Parse(DMA_Array[3].InnerText, CultureInfo.InvariantCulture.NumberFormat);


                    if (DMA_Array[1].InnerText != string.Empty)
                    {
                        ThirtyDMA = float.Parse(DMA_Array[1].InnerText, CultureInfo.InvariantCulture.NumberFormat);
                        FiftyDMA = float.Parse(DMA_Array[3].InnerText, CultureInfo.InvariantCulture.NumberFormat);
                        OneFiftyDMA = float.Parse(DMA_Array[5].InnerText, CultureInfo.InvariantCulture.NumberFormat);
                        TwoHundreadthDMA = float.Parse(DMA_Array[7].InnerText, CultureInfo.InvariantCulture.NumberFormat);

                    }


                }



                worKsheeT.Cells[lastUsedRow + 1, 1] = DateTime.Now;
                worKsheeT.Cells[lastUsedRow + 1, 2] = DateTime.Now.DayOfWeek.ToString();
                worKsheeT.Cells[lastUsedRow + 1, 3] = companyName;
                worKsheeT.Cells[lastUsedRow + 1, 4] = todayOpen;
                worKsheeT.Cells[lastUsedRow + 1, 5] = todayHigh;
                worKsheeT.Cells[lastUsedRow + 1, 6] = todayLow;
                worKsheeT.Cells[lastUsedRow + 1, 7] = todayCurrentPrice;
                worKsheeT.Cells[lastUsedRow + 1, 7].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbOrchid;
                worKsheeT.Cells[lastUsedRow + 1, 8] = previousClose;
                worKsheeT.Cells[lastUsedRow + 1, 9] = ThirtyDMA;
                worKsheeT.Cells[lastUsedRow + 1, 10] = FiftyDMA;
                worKsheeT.Cells[lastUsedRow + 1, 11] = OneFiftyDMA;
                worKsheeT.Cells[lastUsedRow + 1, 12] = TwoHundreadthDMA;
                worKsheeT.Cells[lastUsedRow + 1, 13] = todayAverageVolume;
                worKsheeT.Cells[lastUsedRow + 1, 14] = todayVolume;
                if (i == 0)
                {
                    if (todayVolume != string.Empty && todayAverageVolume != string.Empty)
                    {
                        if (double.Parse(todayVolume) > double.Parse(todayAverageVolume))
                        {
                            worKsheeT.Cells[lastUsedRow + 1, 14].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbYellow;
                        }
                    }


                }

                if (worKsheeT.Cells[lastUsedRow + 1, 4].Value == worKsheeT.Cells[lastUsedRow + 1, 6].Value)
                {
                    worKsheeT.Cells[lastUsedRow + 1, 15].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbDarkGreen;
                }
                else
                {

                    if (worKsheeT.Cells[lastUsedRow + 1, 4].Value == worKsheeT.Cells[lastUsedRow + 1, 5].Value)
                    {
                        worKsheeT.Cells[lastUsedRow + 1, 15].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbDarkRed;
                    }
                }

                if ((worKsheeT.Cells[lastUsedRow + 1, 4].Value > worKsheeT.Cells[lastUsedRow + 1, 6].Value))
                {
                    if ((worKsheeT.Cells[lastUsedRow + 1, 4].Value - worKsheeT.Cells[lastUsedRow + 1, 6].Value) < 0.825)
                    {
                        worKsheeT.Cells[lastUsedRow + 1, 15].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbGreen;
                    }
                }
                else
                {
                    if (worKsheeT.Cells[lastUsedRow + 1, 5].Value > worKsheeT.Cells[lastUsedRow + 1, 4].Value)
                    {
                        if ((worKsheeT.Cells[lastUsedRow + 1, 5].Value - worKsheeT.Cells[lastUsedRow + 1, 4].Value) <= 0.825)
                        {
                            worKsheeT.Cells[lastUsedRow + 1, 15].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbRed;
                        }
                    }
                }

                int x = 0;
                if (todayCurrentPrice > FiftyDMA)
                {
                    if (todayCurrentPrice < 100000)
                    {
                        if (todayCurrentPrice > 300)
                        {
                            if ((todayCurrentPrice - FiftyDMA) <= 5 && (todayCurrentPrice - FiftyDMA) >= 0)
                            {
                                worKsheeT.Cells[lastUsedRow + 1, 16].Value = "Stock is nearer to 50DMA, it may take support.";
                                worKsheeT.Cells[lastUsedRow + 1, 16].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbGreen;
                                x = 1;
                            }

                        }

                        else if (todayCurrentPrice < 300 && todayCurrentPrice > 200)
                        {
                            if ((todayCurrentPrice - FiftyDMA) <= 3 && (todayCurrentPrice - FiftyDMA) >= 0)
                            {
                                worKsheeT.Cells[lastUsedRow + 1, 16].Value = "200 to 300 range Stock is nearer to 50DMA, it may take support";
                                worKsheeT.Cells[lastUsedRow + 1, 16].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbGreen;
                                x = 1;
                            }

                        }

                        else if (todayCurrentPrice < 200 && todayCurrentPrice > 100)
                        {
                            if ((todayCurrentPrice - FiftyDMA) <= 2 && (todayCurrentPrice - FiftyDMA) >= 0)
                            {
                                worKsheeT.Cells[lastUsedRow + 1, 16].Value = "100 to 200 range Stock is nearer to 50DMA, it may take support";
                                worKsheeT.Cells[lastUsedRow + 1, 16].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbGreen;
                                x = 1;
                            }

                        }
                        else if (todayCurrentPrice < 100 && todayCurrentPrice > 50)
                        {
                            if ((todayCurrentPrice - FiftyDMA) <= 1.5 && (todayCurrentPrice - FiftyDMA) >= 0)
                            {
                                worKsheeT.Cells[lastUsedRow + 1, 16].Value = "50 to 100 range Stock is nearer to 50DMA, it may take support";
                                worKsheeT.Cells[lastUsedRow + 1, 16].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbDarkGreen;
                                x = 1;
                            }

                        }
                        else if (todayCurrentPrice < 50)
                        {
                            if ((todayCurrentPrice - FiftyDMA) <= 1 && (todayCurrentPrice - FiftyDMA) >= 0)
                            {
                                worKsheeT.Cells[lastUsedRow + 1, 16].Value = "Below range Stock is nearer to 50DMA, it may take support";
                                worKsheeT.Cells[lastUsedRow + 1, 16].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbDarkGreen;
                                x = 1;
                            }

                        }

                        if (x == 0)
                        {
                            if (todayCurrentPrice > FiftyDMA)
                            {
                                if (todayCurrentPrice > ThirtyDMA)
                                {
                                    worKsheeT.Cells[lastUsedRow + 1, 16].Value = "Stock is trading above 3ODMA and below 50 DMA might be";
                                    worKsheeT.Cells[lastUsedRow + 1, 16].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbLightYellow;
                                }
                                else if (todayCurrentPrice < ThirtyDMA)
                                {
                                    worKsheeT.Cells[lastUsedRow + 1, 16].Value = "Stock is trading below 30DMA and below 50DMA might be. Weakness coming mifgt be";
                                    worKsheeT.Cells[lastUsedRow + 1, 16].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbPaleVioletRed;
                                }
                            }
                            else if ((todayCurrentPrice < FiftyDMA))
                            {
                                worKsheeT.Cells[lastUsedRow + 1, 16].Value = "Stock is closing below 50 DMA it may go down";
                                worKsheeT.Cells[lastUsedRow + 1, 16].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbDarkRed;

                            }

                        }

                    }
                }
                else if (todayCurrentPrice < FiftyDMA)
                {
                    if (todayCurrentPrice < 100000)
                    {
                        if (todayCurrentPrice > 300)
                        {
                            if ((todayCurrentPrice - TwoHundreadthDMA) <= 5 && (todayCurrentPrice - TwoHundreadthDMA) >= 0)
                            {
                                worKsheeT.Cells[lastUsedRow + 1, 16].Value = "Bigger Stock is nearer to 200DMA, it may take support";
                                worKsheeT.Cells[lastUsedRow + 1, 16].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbGreen;
                                x = 1;
                            }

                        }

                        else if (todayCurrentPrice < 300 && todayCurrentPrice > 200 && (todayCurrentPrice - TwoHundreadthDMA) >= 0)
                        {
                            if ((todayCurrentPrice - TwoHundreadthDMA) <= 3)
                            {
                                worKsheeT.Cells[lastUsedRow + 1, 16].Value = "200 to 300 range Stock is nearer to 200DMA, it may take support";
                                worKsheeT.Cells[lastUsedRow + 1, 16].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbGreen;
                                x = 1;
                            }

                        }

                        else if (todayCurrentPrice < 200 && todayCurrentPrice > 100 && (todayCurrentPrice - TwoHundreadthDMA) >= 0)
                        {
                            if ((todayCurrentPrice - TwoHundreadthDMA) <= 2)
                            {
                                worKsheeT.Cells[lastUsedRow + 1, 16].Value = "100 to 200 range Stock is nearer to 200DMA, it may take support";
                                worKsheeT.Cells[lastUsedRow + 1, 16].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbGreen;
                                x = 1;
                            }

                        }
                        else if (todayCurrentPrice < 100 && todayCurrentPrice > 50)
                        {
                            if ((todayCurrentPrice - TwoHundreadthDMA) <= 1.5 && (todayCurrentPrice - TwoHundreadthDMA) >= 0)
                            {
                                worKsheeT.Cells[lastUsedRow + 1, 16].Value = "50 to 100 range Stock is nearer to 200DMA, it may take support";
                                worKsheeT.Cells[lastUsedRow + 1, 16].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbGreen;
                                x = 1;
                            }

                        }
                        else if (todayCurrentPrice < 50)
                        {
                            if ((todayCurrentPrice - TwoHundreadthDMA) <= 1 && (todayCurrentPrice - TwoHundreadthDMA) >= 0)
                            {
                                worKsheeT.Cells[lastUsedRow + 1, 16].Value = "Below 50 range Stock is nearer to 200DMA, it may take support";
                                worKsheeT.Cells[lastUsedRow + 1, 16].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbGreen;
                                x = 1;
                            }

                        }

                        if (x == 0)
                        {
                            if (todayCurrentPrice > TwoHundreadthDMA)
                            {
                                if (todayCurrentPrice > ThirtyDMA)
                                {
                                    worKsheeT.Cells[lastUsedRow + 1, 16].Value = "Stock is trading above 30DMA but below 50DMA might be..";
                                    worKsheeT.Cells[lastUsedRow + 1, 16].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbDarkGreen;
                                }
                                else if (todayCurrentPrice < ThirtyDMA)
                                {
                                    worKsheeT.Cells[lastUsedRow + 1, 16].Value = "Stock is trading below 30DMA and might be below 50DMA. Weakness coming migh t be";
                                    worKsheeT.Cells[lastUsedRow + 1, 16].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbPaleVioletRed;
                                }
                            }
                            else if ((todayCurrentPrice < TwoHundreadthDMA))
                            {
                                worKsheeT.Cells[lastUsedRow + 1, 16].Value = "Stock is closing below 200 DMA it may go down";
                                worKsheeT.Cells[lastUsedRow + 1, 16].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbDarkRed;
                            }
                        }
                    }
                }

                else if (todayCurrentPrice < TwoHundreadthDMA)
                {
                    worKsheeT.Cells[lastUsedRow + 1, 16].Value = "Stock is closing below 200 DMA it may go down";
                    worKsheeT.Cells[lastUsedRow + 1, 16].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbDarkRed;
                }
                else if (todayCurrentPrice < TwoHundreadthDMA)
                {
                    worKsheeT.Cells[lastUsedRow + 1, 16].Value = "Stock is closing below 50 DMA it may go down";
                    worKsheeT.Cells[lastUsedRow + 1, 16].Interior.Color = Microsoft.Office.Interop.Excel.XlRgbColor.rgbDarkRed;

                }
                lastUsedRow = lastUsedRow + 1;

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message.ToString());

            }



        }
    }
}
