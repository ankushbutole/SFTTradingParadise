using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Linq;

namespace NewsTracker
{
 
    class Program
    {
        private static HtmlWeb webNifty;

        static void Main(string[] args)
        {
            //Console.WriteLine("Hello World-!" , DateTime.Now);
              
            webNifty = new HtmlWeb();
           // webBankNifty = new HtmlWeb();
            var niftyUrlvalue = @"https://economictimes.indiatimes.com/markets/stocks?from=mdr";
            //var bankniftyUrlvalue = @"https://www.moneycontrol.com/indian-indices/bank-nifty-23.html";

            HtmlDocument loadniftyhtml = webNifty.Load(niftyUrlvalue);
            loadniftyhtml.LoadHtml(loadniftyhtml.DocumentNode.InnerHtml.ToString());

            //var ecotimesreccmondation = loadniftyhtml.GetElementbyId("bseid").GetAttributeValue("value", "");

            HtmlNode[] todayCurrentPrice_Array = loadniftyhtml.DocumentNode.SelectNodes("//ul[@class='list7']").ToArray();
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Recommendation from brokerage firms , do your own research to trade.", DateTime.Today);
            Console.ForegroundColor = ConsoleColor.Green;
            //Console.WriteLine();
            int tabSize = 8;
            WriteLineWordWrap(todayCurrentPrice_Array[0].InnerText,tabSize);
            Console.ReadKey();

        }

        public static void WriteLineWordWrap(string paragraph, int tabSize = 8)
        {
            string[] lines = paragraph
                .Replace("\t", new String(' ', tabSize))
                .Split(new string[] { Environment.NewLine }, StringSplitOptions.None);

            for (int i = 0; i < lines.Length; i++)
            {
                string process = lines[i];
                List<String> wrapped = new List<string>();

                while (process.Length > Console.WindowWidth)
                {
                    int wrapAt = process.LastIndexOf(' ', Math.Min(Console.WindowWidth - 1, process.Length));
                    if (wrapAt <= 0) break;

                    wrapped.Add(process.Substring(0, wrapAt));
                    process = process.Remove(0, wrapAt + 1);
                }

                foreach (string wrap in wrapped)
                {
                    Console.WriteLine(wrap);
                }

                Console.WriteLine(process);
            }
        }
    }
}
