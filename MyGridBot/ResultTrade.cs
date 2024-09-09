using Bybit.Net.Clients;
using Bybit.Net.Interfaces.Clients;
using ClosedXML.Excel;
using CryptoExchange.Net.Authentication;
using CryptoExchange.Net.CommonObjects;
using CryptoExchange.Net.Objects;
using DocumentFormat.OpenXml.Spreadsheet;
using Mexc.Net.Clients;
using Mexc.Net.Objects.Models.Spot;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyGridBot
{
    internal class ResultTrade
    {
        public static ulong Buy { get; set; } = 0;
        public static ulong Sell { get; set; } = 0;
        static decimal TotalBalanceUSDT { get; set; } = 0;
        static decimal TotalBalanceUSDC { get; set; } = 0;
        static decimal ExpectedProfitUSDT { get; set; } = 0;
        static decimal ExpectedProfitUSDC { get; set; } = 0;
        static int Copy { get; set; } = 0;

        public static async Task BalanceByBit(BybitRestClient bybitRestClient, DateTime dateTime)
        {
            Copy++;
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.WriteLine(" Делаю запрос баланса");
            WebCallResult<Bybit.Net.Objects.Models.V5.BybitResponse<Bybit.Net.Objects.Models.V5.BybitBalance>> balance = null;
            while (true)
            {
                try { balance = await bybitRestClient.V5Api.Account.GetBalancesAsync(Bybit.Net.Enums.AccountType.Unified); }
                catch (Exception ex) { Console.WriteLine(ex.Message); await Task.Delay(1000); }
                if (balance.Error == null) { break; }
                else if (balance.Error.Code == 10002)
                {
                    Console.WriteLine($" {balance.Error.Code} {balance.Error.Message}");
                    await Task.Delay(2000);
                    continue;
                }
                else if (balance.Error.Code == 502)
                {
                    Console.WriteLine($" {balance.Error.Code} {balance.Error.Message}");
                    await Task.Delay(2000);
                    continue;
                }
                else if (balance.Error.Code != null && balance.Error.Message != null)
                {
                    Console.WriteLine($" Ошибка при запросе баланса \n" +
                                      $" {balance.Error.Code} {balance.Error.Message}");
                    Console.ReadLine();
                }
                else
                {
                    await Task.Delay(2000);
                }
            }

            TotalBalanceUSDT = 0;
            TotalBalanceUSDC = 0;
            ExpectedProfitUSDT = 0;
            ExpectedProfitUSDC = 0;

            Console.WriteLine();
            var dt = DateTime.Now;
            var timeElapsed = dt - dateTime;
            Console.WriteLine($" Время работы: \n" +
                              $" Дни: {timeElapsed.Days}  {timeElapsed.Hours:00}:{timeElapsed.Minutes:00}:{timeElapsed.Seconds:00}");
            foreach (var Symbol in SettingStart.SymbolList)
            {
                //USDT
                string asset = "";
                for (int i = 0; i < Symbol.Length - 4; i++)
                {
                    asset += Symbol[i];
                }
                while (true)
                {
                    try
                    {
                        using (var workbook = new XLWorkbook($@"..\\..\\..\\..\\Work\\{Symbol}.xlsx"))
                        {
                            var sheet = workbook.Worksheet(1);

                            if ('T' == Symbol.Last())
                            {
                                TotalBalanceUSDT += Convert.ToDecimal(sheet.Cell(1, 12).Value);
                                ExpectedProfitUSDT += Convert.ToDecimal(sheet.Cell(1, 18).Value);
                            }
                            else
                            {
                                TotalBalanceUSDC += Convert.ToDecimal(sheet.Cell(1, 12).Value);
                                ExpectedProfitUSDC += Convert.ToDecimal(sheet.Cell(1, 18).Value);
                            }

                            foreach (var coins in balance.Data.List)
                            {
                                foreach (var coin in coins.Assets)
                                {
                                    if (coin.Asset == asset)
                                    {
                                        if (coin.WalletBalance < Convert.ToDecimal(sheet.Cell(1, 13).Value))
                                        {
                                            Console.WriteLine($" Монета:{asset} меньше в наличии чем в ексель,на {Convert.ToDecimal(sheet.Cell(1, 13).Value) - coin.WalletBalance}");
                                            Console.ReadLine();
                                        }
                                        else
                                        {
                                            Console.Write($" Монета: ");
                                            Console.ForegroundColor = ConsoleColor.Gray;
                                            Console.Write($"{asset} ");
                                            Console.ForegroundColor = ConsoleColor.Magenta;
                                            Console.Write(" Профит: ");
                                            Console.ForegroundColor = ConsoleColor.Yellow;
                                            Console.Write($"{coin.WalletBalance - Convert.ToDecimal(sheet.Cell(1, 13).Value)}");
                                            Console.ForegroundColor = ConsoleColor.Magenta;
                                            Console.WriteLine();
                                            break;
                                        }
                                    }
                                }

                            }
                        }
                        break;
                    }
                    catch
                    {
                        Console.WriteLine($" Не смог открыть файл {Symbol}.xlsx метод Balance");
                        Thread.Sleep(10000);
                    }
                }
            }
            foreach (var coins in balance.Data.List)
            {
                foreach (var coin in coins.Assets)
                {
                    if (coin.Asset == "USDT")
                    {
                        if (coin.WalletBalance < TotalBalanceUSDT)
                        {
                            Console.WriteLine($" USDT на счете меньше чем нужно для сетки,на {TotalBalanceUSDT - coin.WalletBalance}");
                            Console.ReadLine();
                        }
                        else
                        {
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.Write($" USDT");
                            Console.ForegroundColor = ConsoleColor.Magenta;
                            Console.Write(" Профит: ");
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.Write($"{coin.WalletBalance - TotalBalanceUSDT} $");
                            Console.ForegroundColor = ConsoleColor.Magenta;
                            Console.WriteLine();

                            Console.ForegroundColor = ConsoleColor.Magenta;
                            Console.Write(" Ожидаемая стомость портфеля: ");
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.Write($"{coin.WalletBalance + ExpectedProfitUSDT} $");
                            Console.ForegroundColor = ConsoleColor.Magenta;
                            Console.WriteLine();
                        }
                    }
                    if (coin.Asset == "USDC")
                    {
                        if (coin.WalletBalance < TotalBalanceUSDC)
                        {
                            Console.WriteLine($" USDC на счете меньше чем нужно для сетки,на {TotalBalanceUSDC - coin.WalletBalance}");
                            Console.ReadLine();
                        }
                        else
                        {
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.Write($" USDC");
                            Console.ForegroundColor = ConsoleColor.Magenta;
                            Console.Write(" Профит: ");
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.Write($"{coin.WalletBalance - TotalBalanceUSDC} $");
                            Console.ForegroundColor = ConsoleColor.Magenta;
                            Console.WriteLine();

                            Console.ForegroundColor = ConsoleColor.Magenta;
                            Console.Write(" Ожидаемая стомость портфеля: ");
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.Write($"{coin.WalletBalance + ExpectedProfitUSDC} $");
                            Console.ForegroundColor = ConsoleColor.Magenta;
                            Console.WriteLine();
                        }
                    }
                }
            }
            Console.WriteLine();
            Console.Write($" Сделки: Buy: ");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.Write($"{Buy}");
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.Write(" Sell: ");
            Console.ForegroundColor = ConsoleColor.Red;
            Console.Write($"{Sell}");
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.WriteLine();

            if (50 - Copy > 0)
            {
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine();
                Console.Write(" Сортировка будет через (");
                Console.ForegroundColor = ConsoleColor.White;
                Console.Write($"{50 - Copy}");
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine(") прокрутов");
                Console.ForegroundColor = ConsoleColor.Magenta;
            }

            if (Copy >= 50)
            {
                //Сортировка
                Console.ForegroundColor = ConsoleColor.Blue;
                await NewExcel.SortBuySellByBitAsync(bybitRestClient);
                Copy = 0;
                CopyTable.Copy(@"..\\..\\..\\..\\Work", @"..\\..\\..\\..\\WorkCopy");
            }
        }
        public static async Task BalanceMexc(MexcRestClient mexcRestClient, DateTime dateTime)
        {
            Copy++;
            Console.WriteLine();
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.WriteLine(" Делаю запрос баланса");
            WebCallResult<MexcAccountInfo> balance = null;
            while (true)
            {
                try { balance = await mexcRestClient.SpotApi.Account.GetAccountInfoAsync(); }
                catch (Exception ex) { Console.WriteLine(ex.Message); await Task.Delay(1000); }
                if (balance.Error == null) { break; }
                else if(balance.Error.Code == 700022) { continue; }
                else if (balance.Error.Code == 700003) { continue; }
                else if (balance.Error.Code == 504) { await Task.Delay(3000); continue; }
                else if (balance.Error.Code != null && balance.Error.Message != null)
                {
                    Console.WriteLine($" Ошибка при запросе баланса \n" +
                                      $" {balance.Error.Code} {balance.Error.Message}");
                    Console.ReadLine();
                }
                else
                {
                    await Task.Delay(2000);
                }
            }

            TotalBalanceUSDT = 0;
            TotalBalanceUSDC = 0;
            ExpectedProfitUSDT = 0;
            ExpectedProfitUSDC = 0;

            Console.WriteLine();
            var dt = DateTime.Now;
            var timeElapsed = dt - dateTime;
            Console.WriteLine($" Время работы: \n" +
                              $" Дни: {timeElapsed.Days}  {timeElapsed.Hours:00}:{timeElapsed.Minutes:00}:{timeElapsed.Seconds:00}");
            foreach (var Symbol in SettingStart.SymbolList)
            {
                //USDT
                string asset = "";
                for (int i = 0; i < Symbol.Length - 4; i++)
                {
                    asset += Symbol[i];
                }
                while (true)
                {
                    try
                    {
                        using (var workbook = new XLWorkbook($@"..\\..\\..\\..\\WorkMexc\\{Symbol}.xlsx"))
                        {
                            var sheet = workbook.Worksheet(1);

                            if ('T' == Symbol.Last())
                            {
                                TotalBalanceUSDT += Convert.ToDecimal(sheet.Cell(1, 12).Value);
                                ExpectedProfitUSDT += Convert.ToDecimal(sheet.Cell(1, 18).Value);
                            }
                            else
                            {
                                TotalBalanceUSDC += Convert.ToDecimal(sheet.Cell(1, 12).Value);
                                ExpectedProfitUSDC += Convert.ToDecimal(sheet.Cell(1, 18).Value);
                            }

                            foreach (var coin in balance.Data.Balances)
                            {
                                if (coin.Asset == asset)
                                {
                                    if (coin.Total < Convert.ToDecimal(sheet.Cell(1, 13).Value))
                                    {
                                        Console.WriteLine($" Монета:{asset} меньше в наличии чем в ексель,на {Convert.ToDecimal(sheet.Cell(1, 13).Value) - coin.Total}");
                                        Console.ReadLine();
                                    }
                                    else
                                    {
                                        Console.Write($" Монета: ");
                                        Console.ForegroundColor = ConsoleColor.Gray;
                                        Console.Write($"{asset} ");
                                        Console.ForegroundColor = ConsoleColor.Magenta;
                                        Console.Write(" Профит: ");
                                        Console.ForegroundColor = ConsoleColor.Yellow;
                                        Console.Write($"{coin.Total - Convert.ToDecimal(sheet.Cell(1, 13).Value)}");
                                        Console.ForegroundColor = ConsoleColor.Magenta;
                                        Console.WriteLine();
                                        break;
                                    }
                                }
                            }
                        }
                        break;
                    }
                    catch
                    {
                        Console.WriteLine($" Не смог открыть файл {Symbol}.xlsx метод BalanceMexc");
                        Thread.Sleep(10000);
                    }
                }
            }
            foreach (var coin in balance.Data.Balances)
            {
                if (coin.Asset == "USDT")
                {
                    if (coin.Total < TotalBalanceUSDT)
                    {
                        Console.WriteLine($" USDT на счете меньше чем нужно для сетки,на {TotalBalanceUSDT - coin.Total}");
                        Console.ReadLine();
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.Write($" USDT");
                        Console.ForegroundColor = ConsoleColor.Magenta;
                        Console.Write(" Профит: ");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.Write($"{coin.Total - TotalBalanceUSDT} $");
                        Console.ForegroundColor = ConsoleColor.Magenta;
                        Console.WriteLine();

                        Console.ForegroundColor = ConsoleColor.Magenta;
                        Console.Write(" Ожидаемая стомость портфеля: ");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.Write($"{coin.Total + ExpectedProfitUSDT} $");
                        Console.ForegroundColor = ConsoleColor.Magenta;
                        Console.WriteLine();
                    }
                }
                if (coin.Asset == "USDC")
                {
                    if (coin.Total < TotalBalanceUSDC)
                    {
                        Console.WriteLine($" USDC на счете меньше чем нужно для сетки,на {TotalBalanceUSDC - coin.Total}");
                        Console.ReadLine();
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.Write($" USDC");
                        Console.ForegroundColor = ConsoleColor.Magenta;
                        Console.Write(" Профит: ");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.Write($"{coin.Total - TotalBalanceUSDC} $");
                        Console.ForegroundColor = ConsoleColor.Magenta;
                        Console.WriteLine();

                        Console.ForegroundColor = ConsoleColor.Magenta;
                        Console.Write(" Ожидаемая стомость портфеля: ");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.Write($"{coin.Total + ExpectedProfitUSDC} $");
                        Console.ForegroundColor = ConsoleColor.Magenta;
                        Console.WriteLine();
                    }
                }

            }
            Console.WriteLine();
            Console.Write($" Сделки: Buy: ");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.Write($"{Buy}");
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.Write(" Sell: ");
            Console.ForegroundColor = ConsoleColor.Red;
            Console.Write($"{Sell}");
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.WriteLine();

            if (50 - Copy > 0)
            {
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine();
                Console.Write(" Сортировка будет через (");
                Console.ForegroundColor = ConsoleColor.White;
                Console.Write($"{50 - Copy}");
                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine(") прокрутов");
                Console.ForegroundColor = ConsoleColor.Magenta;
            }

            if (Copy >= 50)
            {
                await NewExcel.SortBuySellMexcAsync();
                Copy = 0;
                CopyTable.Copy(@"..\\..\\..\\..\\WorkMexc", @"..\\..\\..\\..\\WorkCopyMexc");
            }
        }
        public static async Task TimerReversAsync(int seconds, BybitRestClient bybitRestClient)
        {
            bool isPaused = false;
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine();
            Console.WriteLine(" Нажмите ПРОБЕЛ для остановки\n Что бы редактировать все ексель файлы");
            var dateTime = DateTime.Now;
            DateTime dt = dateTime.AddSeconds(-seconds);

            ConsoleKeyInfo keyInfo;
            while (dateTime >= dt)
            {
                if (Console.KeyAvailable)
                {
                    keyInfo = Console.ReadKey(true);
                    if (keyInfo.Key == ConsoleKey.Spacebar)
                    {
                        isPaused = !isPaused;
                        Console.WriteLine(isPaused ? " Таймер приостановлен. \n Можно редактировать ексель\n Нажмите ПРОБЕЛ для продолжения." : " Таймер продолжает работу.");
                        if (!isPaused)
                        {
                            await SettingStart.StartNewExelAsync(bybitRestClient);
                        }
                        Console.ForegroundColor = ConsoleColor.White;
                    }
                }

                if (!isPaused)
                {
                    var ticks = (dateTime - dt).Ticks;
                    Console.WriteLine(new DateTime(ticks).ToString("      HH:mm:ss"));
                    Thread.Sleep(850);
                    dt = dt.AddSeconds(1);
                }
            }
        }
        public static async Task TimerReversAsyncMexc(int seconds, MexcRestClient mexcRestClient)
        {
            bool isPaused = false;
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine();
            Console.WriteLine(" Нажмите ПРОБЕЛ для остановки\n Что бы редактировать все ексель файлы");
            var dateTime = DateTime.Now;
            DateTime dt = dateTime.AddSeconds(-seconds);

            ConsoleKeyInfo keyInfo;
            while (dateTime >= dt)
            {
                if (Console.KeyAvailable)
                {
                    keyInfo = Console.ReadKey(true);
                    if (keyInfo.Key == ConsoleKey.Spacebar)
                    {
                        isPaused = !isPaused;
                        Console.WriteLine(isPaused ? " Таймер приостановлен. \n Можно редактировать ексель\n Нажмите ПРОБЕЛ для продолжения." : " Таймер продолжает работу.");
                        if (!isPaused)
                        {
                            await SettingStart.StartNewExelAsyncMexc(mexcRestClient);
                        }
                        Console.ForegroundColor = ConsoleColor.White;
                    }
                }

                if (!isPaused)
                {
                    var ticks = (dateTime - dt).Ticks;
                    Console.WriteLine(new DateTime(ticks).ToString("      HH:mm:ss"));
                    Thread.Sleep(850);
                    dt = dt.AddSeconds(1);
                }
            }
        }

    }
}
