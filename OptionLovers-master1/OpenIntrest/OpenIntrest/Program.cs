using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Flurl.Http;
using System.Media;
using System.Threading;

namespace OpenIntrest
{
    class Program
    {
        public static string niftyopenintresturl = string.Empty;
        public static string reliance;
        public static string hdfcbank;
        public static float niftyCurrentValue;
        public static float niftyYestrdayLow = 10756;
        public static float niftyYestrdayHigh = 10894;
        public static float niftyYestrdayClose = 10803;

        public static float niftyHourlyHigh = 0;
        public static float niftyHourlyLow = 0;

        public static bool niftyHourlyHighlowstatus = false;

        public static bool niftyHourlyHighbreakstatus = false;
        public static bool niftyHourlylowbreakstatus = false;
        public static bool cancheckhourlydata = false;

        public static bool highbreakstatus = false;
        public static bool closebreakstatus = false;
        public static bool lowbreakStatus = false;
        public static bool niftyhourlyhighSuccess = false;
        public static bool niftyhourlylowsuccess = false;
        // public bool highbreakstatus;

        public static HtmlWeb web;

        public Program()
        {
            niftyopenintresturl = @"https://www1.nseindia.com/live_market/dynaContent/live_watch/option_chain/optionKeys.jsp";
            hdfcbank = "https://www1.nseindia.com/live_market/dynaContent/live_watch/option_chain/optionKeys.jsp?symbolCode=797&symbol=HDFCBANK&symbol=hdfc%20bank&instrument=OPTSTK&date=-&segmentLink=17&segmentLink=17";
            reliance = @"https://www1.nseindia.com/live_market/dynaContent/live_watch/option_chain/optionKeys.jsp?symbolCode=242&symbol=RELIANCE&symbol=Reliance&instrument=OPTSTK&date=-&segmentLink=17&segmentLink=17";
            web = new HtmlWeb();

        }
        //static  void Main1(string[] args)
        //{
        //    Console.WriteLine("Hi");
        //    Console.ReadKey();
        //}

        static async void CallMethod()
        {
            System.Timers.Timer timer = new System.Timers.Timer();
            timer.Interval = 30000;
            timer.Elapsed += timer_Elapsed;
            timer.Start();
        }


        public static void Main(string[] args)
        {
            try
            {
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine("Good morning AB \n First rule of trading is not to loose amount.\n Always goes with Hedge in case of option " +
                    "\n Never buy option in one go.Must be differnce of 4 points." +
                    "\n Trade only two lots strictly.\n Always remember your mistakes and dont Repeat.\n " +
                    "Profitable day should not went into loss\n Remember your goal and how much trading means to you.So one trade can ruin your career \n Follow the charts and system." +
                    "Dimag ki suno Dil ki nai. \n On expiry no trade after 1:00 vrna jo mila vo bhi jayega \n Dont trade on friday infact try to anlyse market.On friday made losses only." +
                    "\n When you want to trade big that must be trending day");

                while (true)
                {
                    Task task = new Task(CallMethod);
                    task.Start();
                    task.Wait();
                    //Console.ReadLine();

                    Console.ReadKey();


                }
               



                //await OpenIntrestAnalysis();
                //var startTimeSpan = TimeSpan.Zero;
                //var periodTimeSpan = TimeSpan.FromMinutes(3);

                //var timer = new System.Threading.Timer(async (e) =>
                //{
                //    try {
                //        await OpenIntrestAnalysis();

                //    }

                //    catch (Exception ex)
                //    {
                //        Console.WriteLine(ex.Message);
                //    }

                //}, null, startTimeSpan, periodTimeSpan);



            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        static async void timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            //YourCode
            await OpenIntrestAnalysis();
        }


        static async Task OpenIntrestAnalysis()
        {
            try
            {
                Program obj = new Program();
                string result = await niftyopenintresturl.GetStringAsync();

                var doc = new HtmlDocument();
                doc.LoadHtml(result);

                var tofindniftyCurrentValue = doc.DocumentNode.SelectNodes("//b");
                string niftycurrentvalueData = tofindniftyCurrentValue[0].InnerText.Replace("NIFTY", "").Trim();
                bool niftycurrentvalueDataSuccess = float.TryParse(niftycurrentvalueData, out niftyCurrentValue);

                if (lowbreakStatus == false)
                {
                    if (niftyCurrentValue <= niftyYestrdayLow)
                    {
                        SystemSounds.Exclamation.Play();
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine(" Nifty has broken yestrday low {0}. Current value of nifty is {1}", niftyYestrdayLow, niftyCurrentValue);
                        lowbreakStatus = true;

                    }

                }
                if ((highbreakstatus == false))
                {
                    if (niftyCurrentValue >= niftyYestrdayHigh)
                    {
                        SystemSounds.Exclamation.Play();
                        Console.ForegroundColor = ConsoleColor.White;
                        Console.WriteLine(" Nifty has broken yestrday high {0}. Current value of nifty is {1}", niftyYestrdayHigh, niftyCurrentValue);
                        highbreakstatus = true;
                    }

                }

                if (niftyCurrentValue >= niftyYestrdayClose)
                {

                    Console.ForegroundColor = ConsoleColor.White;
                    Console.WriteLine(" Nifty is above yestrday close {0}. Current value of nifty is {1}", niftyYestrdayClose, niftyCurrentValue);

                }

                if (niftyCurrentValue <= niftyYestrdayClose)
                {

                    Console.ForegroundColor = ConsoleColor.White;
                    Console.WriteLine(" Nifty is below yestrday close {0}. Current value of nifty is {1} ", niftyYestrdayClose, niftyCurrentValue);

                }

                var totalOIData = doc.DocumentNode.SelectNodes("//table[@id='octable']");
                //var doc1 = new HtmlDocument();
                //doc1.LoadHtml(totalOIData.ToString());
                var totalopenintrestrowdata = doc.DocumentNode.SelectNodes("//tr");
                //HtmlNode node = doc.DocumentNode.SelectSingleNode("//table[@id='octable']/tr[159]");
                //for stock reliance
                //string data= totalopenintrestrowdata[101].InnerText.Trim();
                // for nifty-93 0r 159 or if 103 then 101
                string data = totalopenintrestrowdata[105].InnerText.Trim();

                string[] straingArrayData = data.Split();
                float totalCallOpenIntrest;
                float totalPutOpenIntrest;
                float differnceIncallandputwriting = 0;
                // long data1 = straingArrayData[4].ToString();

                DateTime now = DateTime.Now;

                TimeSpan start = TimeSpan.Parse("10:16");


               


                // TimeSpan now1 = DateTime.Now.TimeOfDay;
                if (niftyHourlyHighlowstatus == false)
                {
                    if (now.TimeOfDay >= start)
                    {
                        try {
                            //match found
                            Console.WriteLine("Please enter first hour candle high");
                            niftyHourlyHigh = Int64.Parse(Console.ReadLine());
                            //niftyhourlyhighSuccess = double.TryParse(Console.ReadLine(), out niftyHourlyHigh);
                            Console.WriteLine("Please enter first hour candle low");
                            niftyHourlyLow = Int64.Parse(Console.ReadLine());
                          
                            niftyHourlyHighlowstatus = true;
                            cancheckhourlydata = true;

                        }

                        catch ( Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                       


                    }
                }

                if (cancheckhourlydata)
                {
                    if (!niftyHourlyHighbreakstatus)
                    {
                        if (niftyCurrentValue >= niftyHourlyHigh)
                        {
                            SystemSounds.Exclamation.Play();
                            Console.ForegroundColor = ConsoleColor.White;
                            Console.WriteLine("Nifty has broken first hour candle high");
                            niftyHourlyHighbreakstatus = true;
                        }

                    }

                }

                if (cancheckhourlydata)
                {
                    if (!niftyHourlylowbreakstatus)
                    {
                        if (niftyCurrentValue <= niftyHourlyLow)
                        {
                            SystemSounds.Exclamation.Play();
                            Console.ForegroundColor = ConsoleColor.White;
                            Console.WriteLine("Nifty has broken first hour candle Low");
                            niftyHourlylowbreakstatus = true;
                        }
                    }


                }


                bool totalCallOpenIntrestSuccess = float.TryParse(straingArrayData[4].ToString().Replace(@",", string.Empty), out totalCallOpenIntrest);
                bool totalPutOpenIntrestSuccess = float.TryParse(straingArrayData[25].ToString().Replace(@",", string.Empty), out totalPutOpenIntrest);
                //bool totalPutOpenIntrestSuccess = Int64.TryParse(straingArrayData[25].ToString(), out totalPutOpenIntrest);

                if (totalCallOpenIntrest > totalPutOpenIntrest)
                {

                    differnceIncallandputwriting = totalCallOpenIntrest - totalPutOpenIntrest;
                    //if (totalCallOpenIntrestSuccess - totalPutOpenIntrest>=3)
                    //{ 

                    //}
                    // Set the Foreground color to blue 
                    Console.ForegroundColor = ConsoleColor.DarkRed;
                    Console.WriteLine("{0} - Call writing is more as compared to put writing by {1}.Resistnace seems increasing at upper level.IF YOU Dumb, GO FOR LONG", now.ToString("F"), differnceIncallandputwriting, Console.ForegroundColor);
                }
                else
                {
                    Console.ForegroundColor
                        = ConsoleColor.DarkGreen;
                    differnceIncallandputwriting = totalPutOpenIntrest - totalCallOpenIntrest;
                    Console.WriteLine("{0} - Put writing is more as compared to call writing by {1}.Supports seems increasing at lower level.IF YOU Dumb, GO FOR SHORT", now.ToString("F"), differnceIncallandputwriting, Console.ForegroundColor);

                }
                //System.Timers.Timer timer = new System.Timers.Timer();
                //timer.Interval = 120000;
                //timer.Elapsed += timer_Elapsed;
                //timer.Start();



                Console.ReadKey();

                //var client = new WebClient();
                //client.Headers.Add("User-Agent", "C# console program");
                //string url = "http://webcode.me";
                //string content = client.DownloadString(niftyopenintresturl);
                //Console.WriteLine(content);

                //Program obj = new Program();
                //niftyopenintresturl = @"https://www1.nseindia.com/live_market/dynaContent/live_watch/option_chain/optionKeys.jsp";
                //HtmlDocument loadfirsthtml = web.Load(niftyopenintresturl);
                //loadfirsthtml.LoadHtml(loadfirsthtml.DocumentNode.InnerHtml.ToString());
                //if (loadfirsthtml.DocumentNode.InnerHtml.ToString() != string.Empty)
                //    await openintrestAnalysis(loadfirsthtml, outRefParams);

            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

    }
}

