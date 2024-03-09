using HtmlAgilityPack;
using System;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Threading;
using System.Threading.Tasks;

namespace NiftyTracker
{
    //public class Out_Ref_Params
    //{
    //    public string stockName;
    //    public string stockTodayClosedPrice;
    //    public string stockTodayStatus;
    //    public string stockRSIValue;
    //    public float MomentumScore;
    //}
    class Program
    {
        private static string[] stockHighLowURLArrayList;
        public static HtmlWeb webNifty;
        public static HtmlWeb webBankNifty;
        public static int  u = 1;
        public static int m = 1;
        public static int j = 1;
        public static int z = 1;
        static void Main(string[] args)
        {
            Console.Title = "SFT Algorithm To Track Indexes";
            Console.WriteLine("Hello Safe Trader! Game of probabilities");
            Console.WriteLine("Nifty script starting at " + DateTime.Now.TimeOfDay);
           
            //var timer = new System.Threading.Timer(
            //e => StockTodayClosedValueAndStatus(),
            //null,
            //TimeSpan.Zero,
            //TimeSpan.FromMinutes(5));


            for (int i = 0; i < 100; i++)
            {
                StockTodayClosedValueAndStatus();
                Thread.Sleep(60 * 1 * 1000);
            }
        }
        public static void StockTodayClosedValueAndStatus()
        {
            string companyName = "";
            try
            {
                webNifty = new HtmlWeb();
                webBankNifty = new HtmlWeb();
                var niftyUrlvalue = @"http://www.moneycontrol.com/indian-indices/cnx-nifty-9.html";
                var bankniftyUrlvalue = @"https://www.moneycontrol.com/indian-indices/bank-nifty-23.html";

                HtmlDocument loadniftyhtml = webNifty.Load(niftyUrlvalue);
                loadniftyhtml.LoadHtml(loadniftyhtml.DocumentNode.InnerHtml.ToString());

                HtmlDocument loadbankniftyhtml = webBankNifty.Load(bankniftyUrlvalue);
                loadbankniftyhtml.LoadHtml(loadbankniftyhtml.DocumentNode.InnerHtml.ToString());

                float niftytodayLow9 = 0;
                float niftytodayHigh9 = 0;
                float bankniftytodayLow9 = 0;
                float bankniftytodayHigh9 = 0;

                float niftytodayLow10 = 0;
                float niftytodayHigh10 = 0;
                float bankniftytodayLow10 = 0;
                float bankniftytodayHigh10 = 0;


                float niftytodayLow2 = 0;
                float niftytodayHigh2 = 0;
                float bankniftytodayLow2 = 0;
                float bankniftytodayHigh2 = 0;

                float niftytodayLowInitial = 0;
                float niftytodayHighInitial = 0;
                float bankniftytodayLowInitial = 0;
                float bankniftytodayHighInitial = 0;

                string todayOpen = string.Empty;
                string todayVolume = string.Empty;
                float todayCurrentPriceNifty = 0;
                float todayCurrentPrice = 0;
                float todayCurrentPriceBankNifty = 0;
                string todayStatusPrice;
                string todayAverageVolume = string.Empty;
                float ThirtyDMA = 0;
                float FiftyDMA = 0;
                float OneFiftyDMA = 0;
                float TwoHundreadthDMA = 0;
                HtmlDocument comapanyNode;
                // string companyName;
                string previousClose;
                if (loadniftyhtml.DocumentNode.InnerHtml.Contains("bseid"))
                {
                    companyName = loadniftyhtml.GetElementbyId("bseid").GetAttributeValue("value", "");
                    //outRefParams.stockName = companyName.ToString();
                    if (companyName == "Nifty")
                    {
                        Console.WriteLine("Nifty");
                    }
                    //When market is not running
                    //HtmlNode[] todayCurrentPrice_Array = loadfirsthtml.DocumentNode.SelectNodes("//span[@class='txt15B nse_span_price_wrap hidden-xs']").ToArray();
                    HtmlNode[] todayCurrentPrice_Array = loadniftyhtml.DocumentNode.SelectNodes("//div[@class='pcstkspr nsestkcp bsestkcp futstkcp optstkcp']").ToArray();


                    todayCurrentPrice = float.Parse(todayCurrentPrice_Array[0].InnerText, CultureInfo.InvariantCulture.NumberFormat);
                    HtmlNode[] todayStatusPrice_Array;
                    //if (loadfirsthtml.DocumentNode.InnerHtml.Contains("nse_span_price_change_prcnt txt14G"))
                    //{
                    //todayStatusPrice_Array = loadniftyhtml.DocumentNode.SelectNodes("//div[@id='stick_ch_prch']").ToArray();
                    //todayStatusPrice = todayStatusPrice_Array[0].InnerText.ToString();
                    //}
                    //else
                    //{
                    //    todayStatusPrice_Array = loadfirsthtml.DocumentNode.SelectNodes("//span[@class='nse_span_price_change_prcnt txt14R hidden-xs']").ToArray();
                    //    todayStatusPrice = todayStatusPrice_Array[0].InnerText.ToString();
                    //}

                }
                else
                {
                    HtmlNode[] todayStatusPrice_Array;
                    var niftytree = loadniftyhtml.GetElementbyId("sp_val");
                    var bankniftytree = loadbankniftyhtml.GetElementbyId("sp_val");

                    todayCurrentPriceNifty = float.Parse(niftytree.InnerText, CultureInfo.InvariantCulture.NumberFormat);
                    todayCurrentPriceBankNifty = float.Parse(bankniftytree.InnerText, CultureInfo.InvariantCulture.NumberFormat);

                    TimeSpan start09 = TimeSpan.Parse("13:13"); // 10 PM
                    TimeSpan end09 = TimeSpan.Parse("13:21");   // 2 AM
                    TimeSpan now09 = DateTime.Now.TimeOfDay;

                  
                    if (u ==1)
                    {
                        var nodeniftytodaylowInitial = loadniftyhtml.GetElementbyId("sp_low");
                        niftytodayLowInitial = float.Parse(nodeniftytodaylowInitial.InnerText, CultureInfo.InvariantCulture.NumberFormat);

                        var nodeniftytodayHighInitial = loadniftyhtml.GetElementbyId("sp_high");
                        niftytodayHighInitial = float.Parse(nodeniftytodayHighInitial.InnerText, CultureInfo.InvariantCulture.NumberFormat);


                        var nodebankniftytodayHighInitial = loadbankniftyhtml.GetElementbyId("sp_high");
                        bankniftytodayHighInitial = float.Parse(nodebankniftytodayHighInitial.InnerText, CultureInfo.InvariantCulture.NumberFormat);
                        var nodebankniftytodaylowInitial = loadbankniftyhtml.GetElementbyId("sp_low");
                        bankniftytodayLowInitial = float.Parse(nodebankniftytodaylowInitial.InnerText, CultureInfo.InvariantCulture.NumberFormat);
                        Console.ForegroundColor = ConsoleColor.DarkYellow; 
                        Console.WriteLine("Nifty high " + " " + niftytodayHighInitial + " " + "Low" + " " + niftytodayLowInitial + " " + "Banknifty high " + " " + bankniftytodayHighInitial + " " + "Low" + " " + bankniftytodayLowInitial + " " + "at" + " " + DateTime.Now.TimeOfDay + " ");
                        highlowdisplay(todayCurrentPriceNifty, todayCurrentPriceBankNifty, niftytodayHighInitial, niftytodayLowInitial, bankniftytodayHighInitial, bankniftytodayLowInitial);
                        Console.ResetColor();
                        u++;

                    }

                    if (start09 <= end09)
                    {
                      
                       
                        // start and stop times are in the same day
                        if (now09 >= start09 && now09 <= end09)
                        {
                            // current time is between start and stop
                            if (m == 1)
                            {
                                var nodeniftytodaylow9 = loadniftyhtml.GetElementbyId("sp_low");
                                niftytodayLow9 = float.Parse(nodeniftytodaylow9.InnerText, CultureInfo.InvariantCulture.NumberFormat);

                                var nodeniftytodayHigh9 = loadniftyhtml.GetElementbyId("sp_high");
                                niftytodayHigh9 = float.Parse(nodeniftytodayHigh9.InnerText, CultureInfo.InvariantCulture.NumberFormat);


                                var nodebankniftytodayHigh9 = loadbankniftyhtml.GetElementbyId("sp_high");
                                bankniftytodayHigh9 = float.Parse(nodebankniftytodayHigh9.InnerText, CultureInfo.InvariantCulture.NumberFormat);
                                var nodebankniftytodaylow9 = loadbankniftyhtml.GetElementbyId("sp_low");
                                bankniftytodayLow9 = float.Parse(nodebankniftytodaylow9.InnerText, CultureInfo.InvariantCulture.NumberFormat);
                                Console.ForegroundColor = ConsoleColor.DarkYellow;
                                Console.WriteLine("Nifty high " + " " + niftytodayHigh9 + " " + "Low" + "" + niftytodayLow9 + " " + "Banknifty high " + " " + bankniftytodayHigh9 + " " + "Low" + bankniftytodayLow9 + " " + "at" + " " + DateTime.Now.TimeOfDay + " ");
                                highlowdisplay(todayCurrentPriceNifty,todayCurrentPriceBankNifty, niftytodayHigh9, niftytodayLow9,bankniftytodayHigh9,bankniftytodayLow9);
                                Console.ResetColor();
                                m++;
                            }
                        }
                    }

                    //int k = 0;
                    //if (k==0)
                    //{
                    //    Console.WriteLine("Nifty high" + " " + niftytodayHigh + " " + "Low" + "" + niftytodayLow +" "+  "Banknifty high " + " " + bankniftytodayHigh + " " + "Low" + bankniftytodayLow + " "+  "at" + " " + DateTime.Now.TimeOfDay + " ");
                    //    k++;
                    //}


                    TimeSpan start10 = TimeSpan.Parse("10:00"); // 10 PM
                    TimeSpan end10 = TimeSpan.Parse("10:15");   // 2 AM
                    TimeSpan now10 = DateTime.Now.TimeOfDay;

                    TimeSpan start2 = TimeSpan.Parse("14:00"); // 10 PM
                    TimeSpan end2 = TimeSpan.Parse("15:30");   // 2 AM
                    TimeSpan now2 = DateTime.Now.TimeOfDay;

                    if (start10 <= end10)
                    {
                        //int j = 1;
                        // start and stop times are in the same day
                        if (now10 >= start10 && now10 <= end10)
                        {
                            // current time is between start and stop
                            if (j == 1)
                            {
                                var nodeniftytodaylow10 = loadniftyhtml.GetElementbyId("sp_low");
                                niftytodayLow10 = float.Parse(nodeniftytodaylow10.InnerText, CultureInfo.InvariantCulture.NumberFormat);

                                var nodeniftytodayHigh10 = loadniftyhtml.GetElementbyId("sp_high");
                                niftytodayHigh10 = float.Parse(nodeniftytodayHigh10.InnerText, CultureInfo.InvariantCulture.NumberFormat);


                                var nodebankniftytodayHigh10 = loadbankniftyhtml.GetElementbyId("sp_high");
                                bankniftytodayHigh10 = float.Parse(nodebankniftytodayHigh10.InnerText, CultureInfo.InvariantCulture.NumberFormat);
                                var nodebankniftytodaylow10 = loadbankniftyhtml.GetElementbyId("sp_low");
                                Console.ForegroundColor = ConsoleColor.DarkYellow;
                                bankniftytodayLow10 = float.Parse(nodebankniftytodaylow10.InnerText, CultureInfo.InvariantCulture.NumberFormat);
                                Console.WriteLine("Nifty high" + " " + niftytodayHigh10 + " " + "Low" + "" + niftytodayLow10 + " " + "Banknifty high " + " " + bankniftytodayHigh10 + " " + "Low" + bankniftytodayLow10 + " " + "at" + " " + DateTime.Now.TimeOfDay + " ");
                                Console.ResetColor();
                                highlowdisplay(todayCurrentPriceNifty, todayCurrentPriceBankNifty, niftytodayHigh10, niftytodayLow10, bankniftytodayHigh10, bankniftytodayLow10);
                                j++;
                            }
                        }
                    }

                    if (start2 <= end2)
                    {
                        // start and stop times are in the same day
                        //int z = 1;
                        if (now2 >= start2 && now2 <= end2)
                        {
                            if (z == 1)
                            {
                                var nodeniftytodaylow2 = loadniftyhtml.GetElementbyId("sp_low");
                                niftytodayLow2 = float.Parse(nodeniftytodaylow2.InnerText, CultureInfo.InvariantCulture.NumberFormat);

                                var nodeniftytodayHigh2 = loadniftyhtml.GetElementbyId("sp_high");
                                niftytodayHigh2 = float.Parse(nodeniftytodayHigh2.InnerText, CultureInfo.InvariantCulture.NumberFormat);


                                var nodebankniftytodayHigh2 = loadbankniftyhtml.GetElementbyId("sp_high");
                                bankniftytodayHigh2 = float.Parse(nodebankniftytodayHigh2.InnerText, CultureInfo.InvariantCulture.NumberFormat);
                                var nodebankniftytodaylow2 = loadbankniftyhtml.GetElementbyId("sp_low");
                                bankniftytodayLow10 = float.Parse(nodebankniftytodaylow2.InnerText, CultureInfo.InvariantCulture.NumberFormat);
                                Console.ForegroundColor = ConsoleColor.DarkYellow;
                                Console.WriteLine("Nifty high" + " " + niftytodayHigh2 + " " + "Low" + "" + niftytodayLow2 + " " + "Banknifty high " + " " + bankniftytodayHigh2 + " " + "Low" + bankniftytodayLow2 + " " + "at" + " " + DateTime.Now.TimeOfDay + " ");
                                Console.ForegroundColor = ConsoleColor.DarkYellow;
                                highlowdisplay(todayCurrentPriceNifty, todayCurrentPriceBankNifty, niftytodayHigh2, niftytodayLow2, bankniftytodayHigh2, bankniftytodayLow2);
                                Console.ResetColor();
                                z++;
                            }  // current time is between start and stop
                        }
                    }

                    //if (todayCurrentPriceNifty > niftytodayHigh)
                    //{
                    //    int t = 0;
                    //    if (t == 0)
                    //    {

                    //    }
                    //}

                    //companyName = loadfirsthtml.GetElementbyId("inid_name FL").GetAttributeValue("value", "");
                    //HtmlNode[] todayCurrentPrice_Array = loadfirsthtml.DocumentNode.SelectNodes("//div[@class='stkdigit']").ToArray();
                    //outRefParams.stockName = companyName.ToString();

                    Console.WriteLine("Nifty current price is" + " " + todayCurrentPriceNifty + " "+ "Banknifty current price is" +" " + todayCurrentPriceBankNifty + " " + "at" + " " + DateTime.Now.TimeOfDay + " ");
                    
                    
                    
                    //todayStatusPrice_Array = loadfirsthtml.DocumentNode.SelectNodes("//div[@id='stick_ch_prch']").ToArray();
                    //todayStatusPrice = todayStatusPrice_Array[0].InnerText.ToString();
                }


                //outRefParams.stockTodayClosedPrice = todayCurrentPrice.ToString();
                //outRefParams.stockTodayStatus = todayStatusPrice.ToString();



                #region traderscockpit
                //RSI
                //web1 = new HtmlWeb();
                //// System.Threading.Thread.Sleep(500);
                //if (companyName == "NIFTY 50")
                //    companyName = "NIFTY";
                //string stockRSILink = @"https://stockholding.traderscockpit.com/?pageView=rsi-indicator-rsi-chart&type=rsi&symbol=" + companyName;
                //HtmlDocument loadfirsthtmlRSIStock = web1.Load(stockRSILink);

                //string stockRSIInspectCode = loadfirsthtmlRSIStock.DocumentNode.InnerHtml;
                //var element = loadfirsthtmlRSIStock.DocumentNode.SelectNodes("//table[@class='greenTable']");

                //HtmlDocument doc = new HtmlDocument();
                //doc.LoadHtml(element[0].InnerHtml.ToString());

                //HtmlNodeCollection nodes = doc.DocumentNode.ChildNodes;

                //var rsiElement = doc.DocumentNode.SelectNodes("//td");

                //outRefParams.stockRSIValue = rsiElement[2].InnerText.ToString();
                ////stockRSIValue = "";
                //Debug.WriteLine("RSI status done for" + companyName);

                #endregion

                //int i = 0;
                //#region "Trendlyn"
                //string link = MainWindow.stockRSIArrayList[index];
                //foreach (string link in MainWindow.stockRSIArrayList)
                //{
                //    if (companyName == "NIFTY 50" || link.Contains("NIFTY50"))
                //    {
                //        outRefParams.stockRSIValue = "";
                //        outRefParams.MomentumScore = 0;
                //    }


                //    else
                //    {
                //        web1 = new HtmlWeb();
                //        // System.Threading.Thread.Sleep(500);

                //        string stockRSILink = link;

                //        //if (stockRSILink != @"https://trendlyne.com/equity/1898/NIFTYBANK/nifty-bank/")
                //        if (false)
                //        {
                //            try
                //            {
                //                HtmlDocument loadfirsthtmlRSIStock = web1.Load(stockRSILink);

                //                string stockRSIInspectCode = loadfirsthtmlRSIStock.DocumentNode.InnerHtml;
                //                var element = loadfirsthtmlRSIStock.DocumentNode.SelectNodes("//table[@class='tl-dataTable table ta-table']");

                //                HtmlDocument doc = new HtmlDocument();
                //                doc.LoadHtml(element[0].InnerHtml.ToString());

                //                HtmlNodeCollection nodes = doc.DocumentNode.ChildNodes;

                //                var rsiElement = doc.DocumentNode.SelectNodes("//td");

                //                outRefParams.stockRSIValue = rsiElement[7].InnerText.Trim().ToString();


                //                //var stockname = loadfirsthtmlRSIStock.DocumentNode.SelectNodes("//p[@class='fs07rem gr']");
                //                //HtmlDocument doc1 = new HtmlDocument();
                //                //doc1.LoadHtml(element[0].InnerHtml.ToString());

                //                //HtmlNodeCollection nodes1 = doc.DocumentNode.ChildNodes;

                //                //var stocknameElemt = doc.DocumentNode.SelectNodes("//p");

                //                //string companyName1 = stocknameElemt[7].InnerText.Trim().ToString();

                //                Debug.WriteLine("RSI status done for" + link + " " + outRefParams.stockRSIValue);

                //                var Momentum_Score = loadfirsthtmlRSIStock.DocumentNode.SelectNodes("//table[@class='tl-dataTable table ta-table']");

                //                HtmlDocument Momentum_Score_Doc = new HtmlDocument();
                //                Momentum_Score_Doc.LoadHtml(Momentum_Score[0].InnerHtml.ToString());

                //                HtmlNodeCollection Momentum_Score_Doc_nodes = Momentum_Score_Doc.DocumentNode.ChildNodes;

                //                var Momentum_Score_Element = Momentum_Score_Doc.DocumentNode.SelectNodes("//td");

                //                outRefParams.MomentumScore = float.Parse(Momentum_Score_Element[1].InnerText.Trim());

                //                Debug.WriteLine("MomentumScore  done for" + link + " " + outRefParams.MomentumScore);
                //            }
                //            catch
                //            {
                //                Debug.WriteLine("RSI status failed for" + link + " " + outRefParams.stockRSIValue);
                //            }

                //            #endregion
                //        }
                //        i++;
                //    }
                //}

                //Debug.WriteLine("Total RSI link count" + i);
            }
            catch (Exception ex)
            {
                // outRefParams.stockTodayClosedPrice = "";
                //outRefParams.stockRSIValue = "";
                //  outRefParams.stockTodayStatus = "";
                //Debug.WriteLine("Exception for" + companyName);

            }



        }

        private static void highlowdisplay(float todayCurrentPriceNifty, float todayCurrentPriceBankNifty, float niftytodayHigh, float niftytodayLow, float bankniftytodayHigh, float bankniftytodayLow)
        {

            if (todayCurrentPriceNifty > niftytodayHigh)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Nifty has broken today high" +  " " + + niftytodayHigh + "at" + " " + DateTime.Now);
                Console.ResetColor();

            }

            if (todayCurrentPriceNifty < niftytodayLow)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Nifty has broken today low" + " " + niftytodayLow + "at" + " " + DateTime.Now);
                Console.ResetColor();

            }

            if (todayCurrentPriceBankNifty > bankniftytodayHigh)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Bank Nifty has broken today high at" + " " + bankniftytodayHigh + "at" + " " + DateTime.Now);
                Console.ResetColor();

            }


            if (todayCurrentPriceBankNifty < bankniftytodayLow)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Bank Nifty has broken today low" + " " + bankniftytodayLow + "at" + " " + DateTime.Now);
                Console.ResetColor();

            }
        }
    }


}


    

