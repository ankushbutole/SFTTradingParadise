using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EquityDailyWPF.Helper
{

    public class Out_Ref_Params
    {
        public string stockName;
        public string stockTodayClosedPrice;
        public string stockTodayStatus;
        public string stockRSIValue;
        public float MomentumScore;
    }
    public static class EquityHelperUtility
    {
        //Stocks
        //string sbin;
        //string pnb;
        //string yesbank;
        //string tcs;
        //string wipro;
        //string infy;
        //string zensartech;
        //string fconsumer;
        //string reliance;
        //string hcc;
        //string itc;
        //string easun;
        //string bjajaauto;
        //string sunpharma;
        //string titan;
        //string bataindia;
        //string cipla;
        //string bajajauto;
        //string hdfc;

        public static HtmlWeb web;
        public static HtmlWeb web1;
        public async static Task StockTodayClosedValueAndStatus(HtmlDocument loadfirsthtml, Out_Ref_Params outRefParams,int index)
        {
            string companyName="";
            try
            {

                string todayLow = string.Empty;
                string todayHigh = string.Empty;
                string todayOpen = string.Empty;
                string todayVolume = string.Empty;
                float todayCurrentPrice;
                string todayStatusPrice;
                string todayAverageVolume = string.Empty;
                float ThirtyDMA = 0;
                float FiftyDMA = 0;
                float OneFiftyDMA = 0;
                float TwoHundreadthDMA = 0;
                HtmlDocument comapanyNode;
               // string companyName;

                string previousClose;
              
                if (loadfirsthtml.DocumentNode.InnerHtml.Contains("bseid"))
                {
                    companyName = loadfirsthtml.GetElementbyId("bseid").GetAttributeValue("value", "");
                    outRefParams.stockName = companyName.ToString();
                    if (companyName == "Nifty")
                    {
                        Console.WriteLine("Nifty");
                    }
                    //When market is not running
                    //HtmlNode[] todayCurrentPrice_Array = loadfirsthtml.DocumentNode.SelectNodes("//span[@class='txt15B nse_span_price_wrap hidden-xs']").ToArray();
                    HtmlNode[] todayCurrentPrice_Array = loadfirsthtml.DocumentNode.SelectNodes("//div[@class='pcstkspr nsestkcp bsestkcp futstkcp optstkcp']").ToArray();


                    todayCurrentPrice = float.Parse(todayCurrentPrice_Array[0].InnerText, CultureInfo.InvariantCulture.NumberFormat);
                    HtmlNode[] todayStatusPrice_Array;
                    //if (loadfirsthtml.DocumentNode.InnerHtml.Contains("nse_span_price_change_prcnt txt14G"))
                    //{
                        todayStatusPrice_Array = loadfirsthtml.DocumentNode.SelectNodes("//div[@id='stick_ch_prch']").ToArray();
                        todayStatusPrice = todayStatusPrice_Array[0].InnerText.ToString();
                    //}
                    //else
                    //{
                    //    todayStatusPrice_Array = loadfirsthtml.DocumentNode.SelectNodes("//span[@class='nse_span_price_change_prcnt txt14R hidden-xs']").ToArray();
                    //    todayStatusPrice = todayStatusPrice_Array[0].InnerText.ToString();
                    //}

                }
                else {
                    HtmlNode[] todayStatusPrice_Array;
                    var tree = loadfirsthtml.GetElementbyId("sp_val");
                    //companyName = loadfirsthtml.GetElementbyId("inid_name FL").GetAttributeValue("value", "");
                    HtmlNode[] todayCurrentPrice_Array = loadfirsthtml.DocumentNode.SelectNodes("//div[@class='stkdigit']").ToArray();
                    outRefParams.stockName = companyName.ToString();
                    todayCurrentPrice = float.Parse(todayCurrentPrice_Array[0].InnerText, CultureInfo.InvariantCulture.NumberFormat);

                    todayStatusPrice_Array = loadfirsthtml.DocumentNode.SelectNodes("//div[@id='stick_ch_prch']").ToArray();
                    todayStatusPrice = todayStatusPrice_Array[0].InnerText.ToString();
                }

              
                outRefParams.stockTodayClosedPrice = todayCurrentPrice.ToString();
                outRefParams.stockTodayStatus = todayStatusPrice.ToString();

                Debug.WriteLine("Closed status done for" + companyName);

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

                int i = 0;
                #region "Trendlyn"
                string link = MainWindow.stockRSIArrayList[index];
                //foreach (string link in MainWindow.stockRSIArrayList)
                {
                    if (companyName == "NIFTY 50" || link.Contains("NIFTY50"))
                    {
                        outRefParams.stockRSIValue = "";
                        outRefParams.MomentumScore = 0;
                    }
                        

                    else {
                        web1 = new HtmlWeb();
                        // System.Threading.Thread.Sleep(500);

                        string stockRSILink = link;

                        //if (stockRSILink != @"https://trendlyne.com/equity/1898/NIFTYBANK/nifty-bank/")
                        if(false)
                        {
                            try {
                                HtmlDocument loadfirsthtmlRSIStock = web1.Load(stockRSILink);

                                string stockRSIInspectCode = loadfirsthtmlRSIStock.DocumentNode.InnerHtml;
                                var element = loadfirsthtmlRSIStock.DocumentNode.SelectNodes("//table[@class='tl-dataTable table ta-table']");

                                HtmlDocument doc = new HtmlDocument();
                                doc.LoadHtml(element[0].InnerHtml.ToString());

                                HtmlNodeCollection nodes = doc.DocumentNode.ChildNodes;

                                var rsiElement = doc.DocumentNode.SelectNodes("//td");

                                outRefParams.stockRSIValue = rsiElement[7].InnerText.Trim().ToString();


                                //var stockname = loadfirsthtmlRSIStock.DocumentNode.SelectNodes("//p[@class='fs07rem gr']");
                                //HtmlDocument doc1 = new HtmlDocument();
                                //doc1.LoadHtml(element[0].InnerHtml.ToString());

                                //HtmlNodeCollection nodes1 = doc.DocumentNode.ChildNodes;

                                //var stocknameElemt = doc.DocumentNode.SelectNodes("//p");

                                //string companyName1 = stocknameElemt[7].InnerText.Trim().ToString();

                                Debug.WriteLine("RSI status done for" + link + " " + outRefParams.stockRSIValue);

                                var Momentum_Score = loadfirsthtmlRSIStock.DocumentNode.SelectNodes("//table[@class='tl-dataTable table ta-table']");

                                HtmlDocument Momentum_Score_Doc = new HtmlDocument();
                                Momentum_Score_Doc.LoadHtml(Momentum_Score[0].InnerHtml.ToString());

                                HtmlNodeCollection Momentum_Score_Doc_nodes = Momentum_Score_Doc.DocumentNode.ChildNodes;

                                var Momentum_Score_Element = Momentum_Score_Doc.DocumentNode.SelectNodes("//td");

                                outRefParams.MomentumScore = float.Parse(Momentum_Score_Element[1].InnerText.Trim());

                                Debug.WriteLine("MomentumScore  done for" + link + " " + outRefParams.MomentumScore);
                            } catch {
                                Debug.WriteLine("RSI status failed for" + link + " " + outRefParams.stockRSIValue);
                            }

                            #endregion
                        }
                        i++;
                    }
                    }

                Debug.WriteLine("Total RSI link count" + i);
            }
            catch (Exception ex)
            {
               // outRefParams.stockTodayClosedPrice = "";
                outRefParams.stockRSIValue = "";
              //  outRefParams.stockTodayStatus = "";
                Debug.WriteLine("Exception for" + companyName);

            }

           

        }

    }
}
