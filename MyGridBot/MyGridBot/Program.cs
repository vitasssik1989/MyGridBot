using Bybit.Net.Clients;
using Bybit.Net.Interfaces.Clients;
using ClosedXML.Excel;
using CryptoExchange.Net.Authentication;
using CryptoExchange.Net.Objects.Options;
using Mexc.Net.Clients;
using Microsoft.Extensions.Options;
using System.Globalization;

namespace MyGridBot
{
    internal class Program
    {
        static async Task Main(string[] args)
        {
            var dateTime = DateTime.Now;
            Console.Title = "BoViGridBot V2.6.6";
            Grafic.GreetUser();
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.WriteLine(" Начинаю работу");
            Console.WriteLine(" На какой бирже хотите работать?\n" +
               " Если BYBIT введите 0 и нажмите ENTER\n" +
               " Если MEXC нажмите 1 и нажмите ENTER");
            if (Console.ReadLine() == "0")
            {
                Console.Title = "BoViGridBot V2.6.6 BYBIT";
                SettingStart.Start();
                BybitRestClient bybitRestClient = new BybitRestClient(options =>
                {
                    options.V5Options.ApiCredentials = new ApiCredentials(SettingStart.APIkey, SettingStart.APIsecret);
                    options.RequestTimeout = TimeSpan.FromSeconds(5);
                    //options.ReceiveWindow = TimeSpan.FromSeconds(30);
                    options.V5Options.TimestampRecalculationInterval = TimeSpan.FromMinutes(1);
                    options.AutoTimestamp = true;
                });

                await SettingStart.StartNewExelAsync(bybitRestClient);
                SettingStart.UpdateSymbolList();

                await ResultTrade.BalanceByBit(bybitRestClient, dateTime);
                while (true)
                {
                    await Trader.BuyByBit(bybitRestClient);
                    await Trader.SellByBit(bybitRestClient);
                    await ResultTrade.BalanceByBit(bybitRestClient, dateTime);
                    await ResultTrade.TimerReversAsync(5, bybitRestClient);
                    SettingStart.UpdateSymbolList();
                }
            }
            else
            {
                Console.Title = "BoViGridBot V2.6.6 MEXC";
                SettingStart.StartMexc();
                MexcRestClient mexcRestClient = new MexcRestClient(opts =>
                {
                    opts.ApiCredentials = new ApiCredentials(SettingStart.APIkey, SettingStart.APIsecret);
                    opts.TimestampRecalculationInterval=TimeSpan.FromMinutes(1);
                    opts.RequestTimeout = TimeSpan.FromSeconds(5);   
                    //opts.ReceiveWindow = TimeSpan.FromSeconds(30);  
                    opts.AutoTimestamp = true;
                });

                await SettingStart.StartNewExelAsyncMexc(mexcRestClient);
                SettingStart.UpdateSymbolListMexc();
                await ResultTrade.BalanceMexc(mexcRestClient, dateTime);

                while (true)
                {
                    await Trader.BuyMexc(mexcRestClient);
                    await Trader.SellMexc(mexcRestClient);
                    await ResultTrade.BalanceMexc(mexcRestClient, dateTime);
                    await ResultTrade.TimerReversAsyncMexc(5, mexcRestClient);
                    SettingStart.UpdateSymbolListMexc();
                }
            }
        }
    }
}