using Bybit.Net.Clients;
using Bybit.Net.Enums;
using Bybit.Net.Objects.Models.Spot;
using Bybit.Net.Objects.Models.V5;
using ClosedXML.Excel;
using CryptoExchange.Net.CommonObjects;
using CryptoExchange.Net.Objects;
using DocumentFormat.OpenXml.Spreadsheet;
using Mexc.Net.Clients;
using Mexc.Net.Objects.Models.Spot;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace MyGridBot
{
    internal class Trader
    {
        public static async Task<BybitOrderbookEntry> AskPriceQuantityByBit(BybitRestClient bybitRestClient, string BuySymbol)
        {
            WebCallResult<BybitOrderbook> orderBookData = null;
            while (true)
            {
                try
                {
                    orderBookData = await bybitRestClient.V5Api.ExchangeData.GetOrderbookAsync(Category.Spot, BuySymbol);

                    if (orderBookData.Error != null)
                    {
                        Console.WriteLine($" Не получил данные по стакану AskPriceQuantity\n" +
                                          $" Ошибка: {orderBookData.Error.Code} {orderBookData.Error.Message}");

                        Thread.Sleep(1000);
                        continue;
                    }
                }
                catch
                {

                    Console.WriteLine($" Не получил данные по стакану AskPriceQuantity\n" +
                                      $" Ошибка: {orderBookData.Error.Code} {orderBookData.Error.Message}");

                    Thread.Sleep(1000);
                }
                if (orderBookData.Data.Asks.First().Price <= 0 && orderBookData.Data.Asks.First().Quantity <= 0)
                {
                    continue;
                }
                break;
            }
            var Ask = orderBookData.Data.Asks.First();
            Console.WriteLine($" AskPrice: {Ask.Price} AskQuantity: {Ask.Quantity}");
            return orderBookData.Data.Asks.First();
        }
        public static async Task<BybitOrderbookEntry> BidPriceQuantityByBit(BybitRestClient bybitRestClient, string SellSymbol)
        {
            WebCallResult<BybitOrderbook> orderBookData = null;
            while (true)
            {
                try
                {
                    orderBookData = await bybitRestClient.V5Api.ExchangeData.GetOrderbookAsync(Category.Spot, SellSymbol);
                    if (orderBookData.Error != null)
                    {
                        Console.WriteLine($" Не получил данные по стакану BidPriceQuantity\n" +
                                          $" Ошибка: {orderBookData.Error.Code} {orderBookData.Error.Message}");

                        Thread.Sleep(1000);
                        continue;
                    }
                }
                catch
                {

                    Console.WriteLine($" Не получил данные по стакану BidPriceQuantity\n" +
                                      $" Ошибка: {orderBookData.Error.Code} {orderBookData.Error.Message}");

                    Thread.Sleep(1000);
                }
                if (orderBookData.Data.Bids.First().Price <= 0 && orderBookData.Data.Bids.First().Quantity <= 0)
                {
                    continue;
                }
                break;
            }
            var Bid = orderBookData.Data.Bids.First();
            Console.WriteLine($" BidPrice: {Bid.Price} BidQuantity: {Bid.Quantity}");
            return orderBookData.Data.Bids.First();
        }

        public static async Task<MexcOrderBookEntry> AskPriceQuantityMexc(MexcRestClient mexcRestClient, string BuySymbol)
        {
            WebCallResult<MexcOrderBook> orderBookData = null;
            while (true)
            {
                try
                {
                    orderBookData = await mexcRestClient.SpotApi.ExchangeData.GetOrderBookAsync(BuySymbol, 1);

                    if (orderBookData.Error != null)
                    {
                        Console.WriteLine($" Не получил данные по стакану AskPriceQuantityMexc\n" +
                                          $" Ошибка: {orderBookData.Error.Code} {orderBookData.Error.Message}");
                        Thread.Sleep(1000);
                        continue;
                    }
                }
                catch
                {

                    Console.WriteLine($" Не получил данные по стакану AskPriceQuantityMexc\n" +
                                      $" Ошибка: {orderBookData.Error.Code} {orderBookData.Error.Message}");
                    Thread.Sleep(1000);
                }
                if (orderBookData.Data.Asks.First().Price <= 0)
                {
                    continue;
                }
                break;
            }
            Console.WriteLine($" AskPrice: {orderBookData.Data.Asks.First().Price} AskQuantity: {orderBookData.Data.Asks.First().Quantity} ");
            return orderBookData.Data.Asks.First();
        }
        public static async Task<MexcOrderBookEntry> BidPriceQuantityMexc(MexcRestClient mexcRestClient, string SellSymbol)
        {
            WebCallResult<MexcOrderBook> orderBookData = null;
            while (true)
            {
                try
                {
                    orderBookData = await mexcRestClient.SpotApi.ExchangeData.GetOrderBookAsync(SellSymbol, 1);
                    if (orderBookData.Error != null)
                    {
                        Console.WriteLine($" Не получил данные по стакану BidPriceQuantityMexc\n" +
                                          $" Ошибка: {orderBookData.Error.Code} {orderBookData.Error.Message}");
                        Thread.Sleep(1000);
                        continue;
                    }
                }
                catch
                {

                    Console.WriteLine($" Не получил данные по стакану BidPriceQuantityMexc\n" +
                                      $" Ошибка: {orderBookData.Error.Code} {orderBookData.Error.Message}");
                    Thread.Sleep(1000);
                }
                if (orderBookData.Data.Bids.First().Price <= 0)
                {
                    continue;
                }
                break;
            }
            Console.WriteLine($" BidPrice: {orderBookData.Data.Bids.First().Price}  BidQuantity: {orderBookData.Data.Bids.First().Quantity}");
            return orderBookData.Data.Bids.First();
        }

        public static async Task BuyByBit(BybitRestClient bybitRestClient)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine();
            Console.WriteLine(" >>>>>>>>>>>> Метод Buy <<<<<<<<<<<<");

            BybitOrderbookEntry Ask = null;

            foreach (var BuySymbol in SettingStart.SymbolList)
            {
                Console.WriteLine();
                Console.WriteLine($" Торговая пара: {BuySymbol}");
                while (true)
                {
                    try
                    {
                        using (var workbook = new XLWorkbook($@"..\\..\\..\\..\\Work\\{BuySymbol}.xlsx"))
                        {
                            bool save = false;
                            var sheet = workbook.Worksheet(1);
                            Ask = await AskPriceQuantityByBit(bybitRestClient, BuySymbol);
                            int grafic = 0;

                            //Трейлинг
                            if (!sheet.Cell(6, 15).IsEmpty())
                            {
                                decimal precent = Convert.ToDecimal(sheet.Cell(6, 15).Value);
                                decimal strategPrice = Convert.ToDecimal(sheet.Cell(4, 15).Value);
                                if (precent > 0)
                                {
                                    if (strategPrice == 0)
                                    {
                                        sheet.Cell(4, 15).Value = Ask.Price;
                                        await Task.Delay(100);
                                        workbook.Save();
                                        Console.WriteLine($" Работает Трейлинг BUY\n Откат {precent} %");
                                        break;
                                    }
                                    else if (strategPrice > Ask.Price)
                                    {
                                        sheet.Cell(4, 15).Value = Ask.Price;
                                        await Task.Delay(100);
                                        workbook.Save();
                                        Console.WriteLine($" Работает Трейлинг BUY\n Откат {precent} %");
                                        break;
                                    }
                                    else if (strategPrice < Ask.Price)
                                    {
                                        if (strategPrice + (strategPrice / 100 * precent) <= Ask.Price)
                                        {
                                            sheet.Cell(4, 15).Value = Ask.Price;
                                            await Task.Delay(100);
                                            workbook.Save();
                                        }
                                        else
                                        {
                                            Console.WriteLine($" Работает Трейлинг BUY\n Откат {precent} %");
                                            break;
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine($" Работает Трейлинг BUY\n Откат {precent} %");
                                        break;
                                    }
                                }
                            }
                            for (int i = 2; i <= 5001; i++)
                            {
                                grafic = i;
                                if (Convert.ToInt32(sheet.Cell(i, 1).Value) == 1 && Convert.ToInt32(sheet.Cell(i, 6).Value) == 1)
                                { continue; }
                                if (Convert.ToInt32(sheet.Cell(i, 1).Value) == 1 && Convert.ToInt32(sheet.Cell(i, 4).Value) == 0)
                                {
                                    if (Ask.Price < Convert.ToDecimal(sheet.Cell(i, 2).Value) && Ask.Quantity > Convert.ToDecimal(sheet.Cell(i, 11).Value))
                                    {
                                        //Реинвестирование
                                        if (Convert.ToInt32(sheet.Cell(i, 5).Value) == 1)
                                        {
                                            Console.WriteLine();
                                            Console.WriteLine($" Покупка Торговой Пары: {BuySymbol}\n" +
                                                              $" По цене: {Convert.ToDecimal(sheet.Cell(i, 2).Value)} \n" +
                                                              $" Кол-во монет: {Convert.ToDecimal(sheet.Cell(i, 11).Value)}\n" +
                                                              $" Реинвестиция: ДА");

                                            if (await BuyResultByBit(bybitRestClient, BuySymbol, Convert.ToDecimal(sheet.Cell(i, 2).Value), Convert.ToDecimal(sheet.Cell(i, 11).Value)))
                                            {
                                                Console.WriteLine(" Заявка исполнилась");
                                                sheet.Cell(i, 8).Value = Convert.ToDecimal(sheet.Cell(i, 11).Value);
                                                sheet.Cell(i, 4).Value = 1;
                                                save = true;
                                                await Task.Delay(200);
                                            }
                                            else
                                            {
                                                Console.WriteLine(" Заявка не исполнилась");
                                                i = 1;
                                                Ask = await AskPriceQuantityByBit(bybitRestClient, BuySymbol);
                                                continue;
                                            }

                                        }
                                        else
                                        {
                                            Console.WriteLine();
                                            Console.WriteLine($" Покупка Торговой Пары: {BuySymbol}\n" +
                                                              $" По цене: {Convert.ToDecimal(sheet.Cell(i, 2).Value)} \n" +
                                                              $" Кол-во монет: {Convert.ToDecimal(sheet.Cell(i, 8).Value)}\n" +
                                                              $" Реинвестиция: НЕТ");

                                            if (await BuyResultByBit(bybitRestClient, BuySymbol, Convert.ToDecimal(sheet.Cell(i, 2).Value), Convert.ToDecimal(sheet.Cell(i, 8).Value)))
                                            {
                                                Console.WriteLine(" Заявка исполнилась");
                                                sheet.Cell(i, 4).Value = 1;
                                                save = true;
                                                await Task.Delay(200);
                                            }
                                            else
                                            {
                                                Console.WriteLine(" Заявка не исполнилась");
                                                i = 1;
                                                Ask = await AskPriceQuantityByBit(bybitRestClient, BuySymbol);
                                                continue;
                                            }

                                        }

                                        await Task.Delay(100);
                                        Ask = await AskPriceQuantityByBit(bybitRestClient, BuySymbol);
                                    }
                                    else
                                    {
                                        Console.WriteLine(" Нет подходящей заявки на покупку");
                                        break;
                                    }
                                }
                                if (i == 5001 && Convert.ToInt32(sheet.Cell(i, 1).Value) == 0)
                                {
                                    Console.BackgroundColor = ConsoleColor.White;
                                    Console.WriteLine(" Закончилась сетка на покупку");
                                    await Task.Delay(1000);
                                    Console.BackgroundColor = ConsoleColor.Black;
                                }
                            }
                            if (save)
                            {
                                workbook.Save();
                            }
                            Grafic.Write(grafic);
                        }
                        break;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($" Не смог открыть файл {BuySymbol}.xlsx");
                        Console.WriteLine(ex.Message); Console.ReadLine();
                    }
                }
            }
        }
        public static async Task SellByBit(BybitRestClient bybitRestClient)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine();
            Console.WriteLine(" >>>>>>>>>>>> Метод Sell <<<<<<<<<<<<");

            BybitOrderbookEntry Bid = null;

            foreach (var SellSymbol in SettingStart.SymbolList)
            {
                Console.WriteLine();
                Console.WriteLine($" Торговая пара: {SellSymbol}");
                while (true)
                {
                    try
                    {
                        using (var workbook = new XLWorkbook($@"..\\..\\..\\..\\Work\\{SellSymbol}.xlsx"))
                        {
                            bool save = false;
                            var sheet = workbook.Worksheet(1);
                            Bid = await BidPriceQuantityByBit(bybitRestClient, SellSymbol);
                            //Трейлинг
                            if (!sheet.Cell(6, 16).IsEmpty())
                            {
                                decimal precent = Convert.ToDecimal(sheet.Cell(6, 16).Value);//0.5
                                decimal strategPrice = Convert.ToDecimal(sheet.Cell(4, 16).Value);//0
                                if (precent > 0)
                                {
                                    if (strategPrice == 0)
                                    {
                                        sheet.Cell(4, 16).Value = Bid.Price;//0.0001245
                                        await Task.Delay(100);
                                        workbook.Save();
                                        Console.WriteLine($" Работает Трейлинг SELL\n Откат {precent} %");
                                        break;
                                    }
                                    else if (strategPrice < Bid.Price)
                                    {
                                        sheet.Cell(4, 16).Value = Bid.Price;
                                        await Task.Delay(100);
                                        workbook.Save();
                                        Console.WriteLine($" Работает Трейлинг SELL\n Откат {precent} %");
                                        break;
                                    }
                                    else if (strategPrice > Bid.Price)
                                    {
                                        if (strategPrice - (strategPrice / 100 * precent) >= Bid.Price)
                                        {
                                            sheet.Cell(4, 16).Value = Bid.Price;
                                            await Task.Delay(100);
                                            workbook.Save();
                                        }
                                        else
                                        {
                                            Console.WriteLine($" Работает Трейлинг SELL\n Откат {precent} %");
                                            break;
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine($" Работает Трейлинг SELL\n Откат {precent} %");
                                        break;
                                    }
                                }
                            }
                            for (int i = 5001; i >= 2; i--)
                            {
                                if (Convert.ToInt32(sheet.Cell(i, 1).Value) == 1 && Convert.ToInt32(sheet.Cell(i, 4).Value) == 1 && Convert.ToInt32(sheet.Cell(i, 6).Value) != 2)
                                {
                                    if (Bid.Price > Convert.ToDecimal(sheet.Cell(i, 3).Value) && Bid.Quantity > Convert.ToDecimal(sheet.Cell(i, 7).Value))
                                    {
                                        Console.WriteLine();
                                        Console.WriteLine($" Продажа Торговой Пары: {SellSymbol}\n" +
                                                              $" По цене: {Convert.ToDecimal(sheet.Cell(i, 3).Value)} \n" +
                                                              $" Кол-во монет: {Convert.ToDecimal(sheet.Cell(i, 7).Value)}");
                                        if (await SellResultByBit(bybitRestClient, SellSymbol, Convert.ToDecimal(sheet.Cell(i, 3).Value), Convert.ToDecimal(sheet.Cell(i, 7).Value)))
                                        {
                                            Console.WriteLine(" Заявка исполнилась");
                                            sheet.Cell(i, 4).Value = 0;
                                            save = true;
                                            await Task.Delay(200);
                                        }
                                        else
                                        {
                                            Console.WriteLine(" Заявка не исполнилась");
                                            i = 5002;
                                            Bid = await BidPriceQuantityByBit(bybitRestClient, SellSymbol);
                                            continue;
                                        }
                                        await Task.Delay(100);
                                        Bid = await BidPriceQuantityByBit(bybitRestClient, SellSymbol);
                                    }
                                    else
                                    {
                                        Console.WriteLine(" Нет подходящей заявки на продажу");
                                        break;
                                    }
                                }
                                if (i == 2 && Convert.ToInt32(sheet.Cell(i, 1).Value) == 0)
                                {
                                    Console.BackgroundColor = ConsoleColor.White;
                                    Console.WriteLine(" Закончилась сетка на продажу");
                                    await Task.Delay(1000);
                                    Console.BackgroundColor = ConsoleColor.Black;
                                }
                            }
                            if (save)
                            {
                                workbook.Save();
                            }
                        }
                        break;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($" Не смог открыть файл {SellSymbol}.xlsx");
                        Console.WriteLine(ex.Message); Console.ReadLine();
                    }
                }

            }
        }

        public static async Task BuyMexc(MexcRestClient mexcRestClient)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine();
            Console.WriteLine(" >>>>>>>>>>>> Метод Buy <<<<<<<<<<<<");

            MexcOrderBookEntry Ask = null;

            foreach (var BuySymbol in SettingStart.SymbolList)
            {
                Console.WriteLine();
                Console.WriteLine($" Торговая пара: {BuySymbol}");
                while (true)
                {
                    try
                    {
                        using (var workbook = new XLWorkbook($@"..\\..\\..\\..\\WorkMexc\\{BuySymbol}.xlsx"))
                        {
                            bool save = false;
                            var sheet = workbook.Worksheet(1);
                            Ask = await AskPriceQuantityMexc(mexcRestClient, BuySymbol);
                            int grafic = 0;

                            //Трейлинг
                            if (!sheet.Cell(6, 15).IsEmpty())
                            {
                                decimal precent = Convert.ToDecimal(sheet.Cell(6, 15).Value);
                                decimal strategPrice = Convert.ToDecimal(sheet.Cell(4, 15).Value);
                                if (precent > 0)
                                {
                                    if (strategPrice == 0)
                                    {
                                        sheet.Cell(4, 15).Value = Ask.Price;
                                        await Task.Delay(100);
                                        workbook.Save();
                                        Console.WriteLine($" Работает Трейлинг BUY\n Откат {precent} %");
                                        break;
                                    }
                                    else if (strategPrice > Ask.Price)
                                    {
                                        sheet.Cell(4, 15).Value = Ask.Price;
                                        await Task.Delay(100);
                                        workbook.Save();
                                        Console.WriteLine($" Работает Трейлинг BUY\n Откат {precent} %");
                                        break;
                                    }
                                    else if (strategPrice < Ask.Price)
                                    {
                                        if (strategPrice + (strategPrice / 100 * precent) <= Ask.Price)
                                        {
                                            sheet.Cell(4, 15).Value = Ask.Price;
                                            await Task.Delay(100);
                                            workbook.Save();
                                        }
                                        else
                                        {
                                            Console.WriteLine($" Работает Трейлинг BUY\n Откат {precent} %");
                                            break;
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine($" Работает Трейлинг BUY\n Откат {precent} %");
                                        break;
                                    }
                                }
                            }
                            for (int i = 2; i <= 5001; i++)
                            {
                                grafic = i;
                                if (Convert.ToInt32(sheet.Cell(i, 1).Value) == 1 && Convert.ToInt32(sheet.Cell(i, 6).Value) == 1)
                                { continue; }
                                if (Convert.ToInt32(sheet.Cell(i, 1).Value) == 1 && Convert.ToInt32(sheet.Cell(i, 4).Value) == 0)
                                {
                                    if (Ask.Price < Convert.ToDecimal(sheet.Cell(i, 2).Value) && Ask.Quantity > Convert.ToDecimal(sheet.Cell(i, 11).Value))
                                    {
                                        //Реинвестирование
                                        if (Convert.ToInt32(sheet.Cell(i, 5).Value) == 1)
                                        {
                                            Console.WriteLine();
                                            Console.WriteLine($" Покупка Торговой Пары: {BuySymbol}\n" +
                                                              $" По цене: {Convert.ToDecimal(sheet.Cell(i, 2).Value)} \n" +
                                                              $" Кол-во монет: {Convert.ToDecimal(sheet.Cell(i, 11).Value)}\n" +
                                                              $" Реинвестиция: ДА");

                                            if (await BuyResultMexc(mexcRestClient, BuySymbol, Convert.ToDecimal(sheet.Cell(i, 2).Value), Convert.ToDecimal(sheet.Cell(i, 11).Value)))
                                            {
                                                Console.WriteLine(" Заявка исполнилась");
                                                sheet.Cell(i, 8).Value = Convert.ToDecimal(sheet.Cell(i, 11).Value);
                                                sheet.Cell(i, 4).Value = 1;
                                                save = true;
                                                await Task.Delay(200);
                                            }
                                            else
                                            {
                                                Console.WriteLine(" Заявка не исполнилась");
                                                i = 1;
                                                Ask = await AskPriceQuantityMexc(mexcRestClient, BuySymbol);
                                                continue;
                                            }

                                        }
                                        else
                                        {
                                            Console.WriteLine();
                                            Console.WriteLine($" Покупка Торговой Пары: {BuySymbol}\n" +
                                                              $" По цене: {Convert.ToDecimal(sheet.Cell(i, 2).Value)} \n" +
                                                              $" Кол-во монет: {Convert.ToDecimal(sheet.Cell(i, 8).Value)}\n" +
                                                              $" Реинвестиция: НЕТ");

                                            if (await BuyResultMexc(mexcRestClient, BuySymbol, Convert.ToDecimal(sheet.Cell(i, 2).Value), Convert.ToDecimal(sheet.Cell(i, 8).Value)))
                                            {
                                                Console.WriteLine(" Заявка исполнилась");
                                                sheet.Cell(i, 4).Value = 1;
                                                save = true;
                                                await Task.Delay(200);
                                            }
                                            else
                                            {
                                                Console.WriteLine(" Заявка не исполнилась");
                                                i = 1;
                                                Ask = await AskPriceQuantityMexc(mexcRestClient, BuySymbol);
                                                continue;
                                            }

                                        }

                                        await Task.Delay(100);
                                        Ask = await AskPriceQuantityMexc(mexcRestClient, BuySymbol);
                                    }
                                    else
                                    {
                                        Console.WriteLine(" Нет подходящей заявки на покупку");
                                        break;
                                    }
                                }
                                if (i == 5001 && Convert.ToInt32(sheet.Cell(i, 1).Value) == 0)
                                {
                                    Console.BackgroundColor = ConsoleColor.White;
                                    Console.WriteLine(" Закончилась сетка на покупку");
                                    await Task.Delay(1000);
                                    Console.BackgroundColor = ConsoleColor.Black;
                                }
                            }
                            if (save)
                            {
                                workbook.Save();
                            }
                            Grafic.Write(grafic);
                        }
                        break;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($" Не смог открыть файл {BuySymbol}.xlsx");
                        Console.WriteLine(ex.Message); Console.ReadLine();
                    }
                }
            }
        }
        public static async Task SellMexc(MexcRestClient mexcRestClient)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine();
            Console.WriteLine(" >>>>>>>>>>>> Метод Sell <<<<<<<<<<<<");

            MexcOrderBookEntry Bid = null;

            foreach (var SellSymbol in SettingStart.SymbolList)
            {
                Console.WriteLine();
                Console.WriteLine($" Торговая пара: {SellSymbol}");
                while (true)
                {
                    try
                    {
                        using (var workbook = new XLWorkbook($@"..\\..\\..\\..\\WorkMexc\\{SellSymbol}.xlsx"))
                        {
                            bool save = false;
                            var sheet = workbook.Worksheet(1);
                            Bid = await BidPriceQuantityMexc(mexcRestClient, SellSymbol);
                            //Трейлинг
                            if (!sheet.Cell(6, 16).IsEmpty())
                            {
                                decimal precent = Convert.ToDecimal(sheet.Cell(6, 16).Value);//0.5
                                decimal strategPrice = Convert.ToDecimal(sheet.Cell(4, 16).Value);//0
                                if (precent > 0)
                                {
                                    if (strategPrice == 0)
                                    {
                                        sheet.Cell(4, 16).Value = Bid.Price;//0.0001245
                                        await Task.Delay(100);
                                        workbook.Save();
                                        Console.WriteLine($" Работает Трейлинг SELL\n Откат {precent} %");
                                        break;
                                    }
                                    else if (strategPrice < Bid.Price)
                                    {
                                        sheet.Cell(4, 16).Value = Bid.Price;
                                        await Task.Delay(100);
                                        workbook.Save();
                                        Console.WriteLine($" Работает Трейлинг SELL\n Откат {precent} %");
                                        break;
                                    }
                                    else if (strategPrice > Bid.Price)
                                    {
                                        if (strategPrice - (strategPrice / 100 * precent) >= Bid.Price)
                                        {
                                            sheet.Cell(4, 16).Value = Bid.Price;
                                            await Task.Delay(100);
                                            workbook.Save();
                                        }
                                        else
                                        {
                                            Console.WriteLine($" Работает Трейлинг SELL\n Откат {precent} %");
                                            break;
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine($" Работает Трейлинг SELL\n Откат {precent} %");
                                        break;
                                    }
                                }
                            }
                            for (int i = 5001; i >= 2; i--)
                            {
                                if (Convert.ToInt32(sheet.Cell(i, 1).Value) == 1 && Convert.ToInt32(sheet.Cell(i, 4).Value) == 1 && Convert.ToInt32(sheet.Cell(i, 6).Value) != 2)
                                {
                                    if (Bid.Price > Convert.ToDecimal(sheet.Cell(i, 3).Value) && Bid.Quantity > Convert.ToDecimal(sheet.Cell(i, 11).Value))
                                    {
                                        Console.WriteLine();
                                        Console.WriteLine($" Продажа Торговой Пары: {SellSymbol}\n" +
                                                              $" По цене: {Convert.ToDecimal(sheet.Cell(i, 3).Value)} \n" +
                                                              $" Кол-во монет: {Convert.ToDecimal(sheet.Cell(i, 7).Value)}");
                                        if (await SellResultMexc(mexcRestClient, SellSymbol, Convert.ToDecimal(sheet.Cell(i, 3).Value), Convert.ToDecimal(sheet.Cell(i, 7).Value)))
                                        {
                                            Console.WriteLine(" Заявка исполнилась");
                                            sheet.Cell(i, 4).Value = 0;
                                            save = true;
                                            await Task.Delay(200);
                                        }
                                        else
                                        {
                                            Console.WriteLine(" Заявка не исполнилась");
                                            i = 5002;
                                            Bid = await BidPriceQuantityMexc(mexcRestClient, SellSymbol);
                                            continue;
                                        }
                                        await Task.Delay(100);
                                        Bid = await BidPriceQuantityMexc(mexcRestClient, SellSymbol);
                                    }
                                    else
                                    {
                                        Console.WriteLine(" Нет подходящей заявки на продажу");
                                        break;
                                    }
                                }
                                if (i == 2 && Convert.ToInt32(sheet.Cell(i, 1).Value) == 0)
                                {
                                    Console.BackgroundColor = ConsoleColor.White;
                                    Console.WriteLine(" Закончилась сетка на продажу");
                                    await Task.Delay(1000);
                                    Console.BackgroundColor = ConsoleColor.Black;
                                }
                            }
                            if (save)
                            {
                                workbook.Save();
                            }
                        }
                        break;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($" Не смог открыть файл {SellSymbol}.xlsx");
                        Console.WriteLine(ex.Message); Console.ReadLine();
                    }
                }

            }
        }

        static async Task<bool> BuyResultByBit(BybitRestClient bybitRestClient, string BuySymbol, decimal price, decimal quantity)
        {
            bool resltBuy = true;
            try
            {
                WebCallResult<Bybit.Net.Objects.Models.V5.BybitOrderId> result = null;
                WebCallResult<Bybit.Net.Objects.Models.V5.BybitResponse<Bybit.Net.Objects.Models.V5.BybitOrder>> resultOrderBuy = null;
                try
                {
                    result = await bybitRestClient.V5Api.Trading.PlaceOrderAsync
                              (
                                  Bybit.Net.Enums.Category.Spot,
                                  symbol: BuySymbol,
                                  side: Bybit.Net.Enums.OrderSide.Buy,
                                  type: Bybit.Net.Enums.NewOrderType.Limit,
                                  price: price,
                                  quantity: quantity,
                                  timeInForce: Bybit.Net.Enums.TimeInForce.FillOrKill
                               );
                }
                catch (Exception ex)
                {
                    Console.WriteLine();
                    if (await CheckExchangeOrderStatusByBit("ByBit", "покупка", BuySymbol, price, quantity, bybitRestClient))
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }

                if (result.Error == null)
                {
                    while (true)
                    {
                        try
                        {
                            resultOrderBuy = await bybitRestClient.V5Api.Trading.GetOrdersAsync
                            (
                                category: Bybit.Net.Enums.Category.Spot,
                                symbol: BuySymbol,
                                clientOrderId: result.Data.ClientOrderId
                            );
                        }
                        catch
                        {
                            Console.WriteLine();
                            if (await CheckExchangeOrderStatusByBit("ByBit", "покупка", BuySymbol, price, quantity, bybitRestClient))
                            {
                                return true;
                            }
                            else
                            {
                                return false;
                            }
                        }

                        if (resultOrderBuy.Error == null)
                        {
                            if (resultOrderBuy.Data.List.First().Status == Bybit.Net.Enums.V5.OrderStatus.Filled)
                            {
                                resltBuy = true;
                                break;
                            }
                            else if (resultOrderBuy.Data.List.First().Status == Bybit.Net.Enums.V5.OrderStatus.Cancelled)
                            {
                                resltBuy = false;
                                break;
                            }
                            await Task.Delay(2000);
                            continue;
                        }
                        else if (resultOrderBuy.Error.Code == 10002 || resultOrderBuy.Error.Code == 170213)
                        {
                            await Task.Delay(1000);
                            continue;
                        }
                        else
                        {
                            Console.WriteLine($" {resultOrderBuy.Error.Code} {resultOrderBuy.Error.Message}");
                            Console.WriteLine();
                            if (await CheckExchangeOrderStatusByBit("ByBit", "покупка", BuySymbol, price, quantity, bybitRestClient))
                            {
                                return true;
                            }
                            else
                            {
                                return false;
                            }
                        }
                    }

                }
                else if (result.Error.Code == 170193)
                {
                    while (true)
                    {
                        BybitOrderbookEntry Ask = await AskPriceQuantityByBit(bybitRestClient, BuySymbol);
                        try
                        {
                            result = await bybitRestClient.V5Api.Trading.PlaceOrderAsync
                                    (
                                       Bybit.Net.Enums.Category.Spot,
                                       symbol: BuySymbol,
                                       side: Bybit.Net.Enums.OrderSide.Buy,
                                       type: Bybit.Net.Enums.NewOrderType.Limit,
                                       price: Ask.Price,
                                       quantity: quantity,
                                       timeInForce: Bybit.Net.Enums.TimeInForce.FillOrKill
                                    );
                        }
                        catch
                        {
                            Console.WriteLine();
                            if (await CheckExchangeOrderStatusByBit("ByBit", "покупка", BuySymbol, Ask.Price, quantity, bybitRestClient))
                            {
                                return true;
                            }
                            else
                            {
                                return false;
                            }
                        }
                        if (result.Error == null)
                        {
                            while (true)
                            {
                                try
                                {

                                    resultOrderBuy = await bybitRestClient.V5Api.Trading.GetOrdersAsync
                                    (
                                        category: Bybit.Net.Enums.Category.Spot,
                                        symbol: BuySymbol,
                                        clientOrderId: result.Data.ClientOrderId
                                    );
                                }
                                catch
                                {
                                    Console.WriteLine();
                                    if (await CheckExchangeOrderStatusByBit("ByBit", "покупка", BuySymbol, Ask.Price, quantity, bybitRestClient))
                                    {
                                        return true;
                                    }
                                    else
                                    {
                                        return false;
                                    }
                                }

                                if (resultOrderBuy.Error == null)
                                {
                                    if (resultOrderBuy.Data.List.First().Status == Bybit.Net.Enums.V5.OrderStatus.Filled)
                                    {
                                        resltBuy = true;
                                        break;
                                    }
                                    else if (resultOrderBuy.Data.List.First().Status == Bybit.Net.Enums.V5.OrderStatus.Cancelled)
                                    {
                                        resltBuy = false;
                                        break;
                                    }
                                    await Task.Delay(1000); continue;
                                }
                                else if (resultOrderBuy.Error.Code == 10002 || resultOrderBuy.Error.Code == 170213)
                                {
                                    await Task.Delay(1000);
                                    continue;
                                }
                                else
                                {
                                    Console.WriteLine($" {resultOrderBuy.Error.Code} {resultOrderBuy.Error.Message}");
                                    Console.WriteLine();
                                    if (CheckExchangeOrderStatus("ByBit", "покупка", BuySymbol, Ask.Price, quantity))
                                    {
                                        return true;
                                    }
                                    else
                                    {
                                        return false;
                                    }
                                }
                            }
                        }
                        else if (result.Error.Code == 10002)
                        {
                            await Task.Delay(2000);
                            continue;
                        }
                        else
                        {
                            Console.WriteLine($" {result.Error.Code} {result.Error.Message}");

                            if (result.Error.Code == 170193)
                            {
                                continue;
                            }
                            Console.WriteLine();
                            if (CheckExchangeOrderStatus("ByBit", "покупка", BuySymbol, Ask.Price, quantity))
                            {
                                return true;
                            }
                            else
                            {
                                return false;
                            }
                        }
                        break;
                    }
                }
                else if (result.Error.Code == 10002)
                {
                    resltBuy = false;
                }
                else
                {
                    Console.WriteLine($" {result.Error.Code} {result.Error.Message}");
                    Console.WriteLine();
                    if (CheckExchangeOrderStatus("ByBit", "покупка", BuySymbol, price, quantity))
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                if (resltBuy == true)
                {
                    ResultTrade.Buy++;
                }
            }
            catch
            {
                Console.WriteLine();
                if (await CheckExchangeOrderStatusByBit("ByBit", "покупка", BuySymbol, price, quantity, bybitRestClient))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            return resltBuy;
        }
        static async Task<bool> SellResultByBit(BybitRestClient bybitRestClient, string SellSymbol, decimal price, decimal quantity)
        {
            bool resltSell = true;
            try
            {
                WebCallResult<Bybit.Net.Objects.Models.V5.BybitOrderId> result = null;
                WebCallResult<Bybit.Net.Objects.Models.V5.BybitResponse<Bybit.Net.Objects.Models.V5.BybitOrder>> resultOrderSell = null;
                try
                {
                    result = await bybitRestClient.V5Api.Trading.PlaceOrderAsync
                               (
                                   Bybit.Net.Enums.Category.Spot,
                                   symbol: SellSymbol,
                                   side: Bybit.Net.Enums.OrderSide.Sell,
                                   type: Bybit.Net.Enums.NewOrderType.Limit,
                                   price: price,
                                   quantity: quantity,
                                   timeInForce: Bybit.Net.Enums.TimeInForce.FillOrKill
                                );
                }
                catch
                {
                    Console.WriteLine();
                    if (await CheckExchangeOrderStatusByBit("ByBit", "продажа", SellSymbol, price, quantity, bybitRestClient))
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }

                if (result.Error == null)
                {
                    while (true)
                    {
                        try
                        {
                            resultOrderSell = await bybitRestClient.V5Api.Trading.GetOrdersAsync
                                (
                                    category: Bybit.Net.Enums.Category.Spot,
                                    symbol: SellSymbol,
                                    clientOrderId: result.Data.ClientOrderId
                                );
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($" {result.Error.Code} {result.Error.Message}");
                            Console.WriteLine();
                            if (await CheckExchangeOrderStatusByBit("ByBit", "продажа", SellSymbol, price, quantity, bybitRestClient))
                            {
                                return true;
                            }
                            else
                            {
                                return false;
                            }
                        }

                        if (resultOrderSell.Error == null)
                        {
                            if (resultOrderSell.Data.List.First().Status == Bybit.Net.Enums.V5.OrderStatus.Filled)
                            {
                                resltSell = true;
                                break;
                            }
                            else if (resultOrderSell.Data.List.First().Status == Bybit.Net.Enums.V5.OrderStatus.Cancelled)
                            {
                                resltSell = false;
                                break;
                            }
                            await Task.Delay(1000); continue;
                        }
                        else if (resultOrderSell.Error.Code == 10002 || resultOrderSell.Error.Code == 170213)
                        {
                            await Task.Delay(1000);
                            continue;
                        }
                        else
                        {
                            Console.WriteLine($" {result.Error.Code} {result.Error.Message}");
                            Console.WriteLine();
                            if (await CheckExchangeOrderStatusByBit("ByBit", "продажа", SellSymbol, price, quantity, bybitRestClient))
                            {
                                return true;
                            }
                            else
                            {
                                return false;
                            }
                        }
                    }

                }
                else if (result.Error.Code == 170194)
                {
                    while (true)
                    {
                        BybitOrderbookEntry Bid = await BidPriceQuantityByBit(bybitRestClient, SellSymbol);
                        try
                        {
                            result = await bybitRestClient.V5Api.Trading.PlaceOrderAsync
                                (
                                    Bybit.Net.Enums.Category.Spot,
                                    symbol: SellSymbol,
                                    side: Bybit.Net.Enums.OrderSide.Sell,
                                    type: Bybit.Net.Enums.NewOrderType.Limit,
                                    price: Bid.Price,
                                    quantity: quantity,
                                    timeInForce: Bybit.Net.Enums.TimeInForce.FillOrKill
                                 );
                        }
                        catch
                        {
                            Console.WriteLine();
                            if (await CheckExchangeOrderStatusByBit("ByBit", "продажа", SellSymbol, Bid.Price, quantity, bybitRestClient))
                            {
                                return true;
                            }
                            else
                            {
                                return false;
                            }
                        }
                        if (result.Error == null)
                        {
                            while (true)
                            {
                                try
                                {
                                    resultOrderSell = await bybitRestClient.V5Api.Trading.GetOrdersAsync
                                    (
                                        category: Bybit.Net.Enums.Category.Spot,
                                        symbol: SellSymbol,
                                        clientOrderId: result.Data.ClientOrderId
                                    );
                                }
                                catch
                                {
                                    Console.WriteLine();
                                    if (await CheckExchangeOrderStatusByBit("ByBit", "продажа", SellSymbol, Bid.Price, quantity, bybitRestClient))
                                    {
                                        return true;
                                    }
                                    else
                                    {
                                        return false;
                                    }
                                }

                                if (resultOrderSell.Error == null)
                                {
                                    if (resultOrderSell.Data.List.First().Status == Bybit.Net.Enums.V5.OrderStatus.Filled)
                                    {
                                        resltSell = true;
                                        break;
                                    }
                                    else if (resultOrderSell.Data.List.First().Status == Bybit.Net.Enums.V5.OrderStatus.Cancelled)
                                    {
                                        resltSell = false;
                                        break;
                                    }
                                    await Task.Delay(1000); continue;
                                }
                                else if (resultOrderSell.Error.Code == 10002 || resultOrderSell.Error.Code == 170213)
                                {
                                    await Task.Delay(1000);
                                    continue;
                                }
                                else
                                {
                                    Console.WriteLine($" {resultOrderSell.Error.Code} {resultOrderSell.Error.Message}");
                                    Console.WriteLine();
                                    if (CheckExchangeOrderStatus("ByBit", "продажа", SellSymbol, Bid.Price, quantity))
                                    {
                                        return true;
                                    }
                                    else
                                    {
                                        return false;
                                    }
                                }
                            }
                        }
                        else if (result.Error.Code == 10002)
                        {
                            await Task.Delay(2000);
                            continue;
                        }
                        else
                        {
                            Console.WriteLine($" {result.Error.Code} {result.Error.Message}");
                            if (result.Error.Code == 170194)
                            {
                                continue;
                            }
                            Console.WriteLine();
                            if (CheckExchangeOrderStatus("ByBit", "продажа", SellSymbol, Bid.Price, quantity))
                            {
                                return true;
                            }
                            else
                            {
                                return false;
                            }
                        }
                        break;
                    }
                }
                else if (result.Error.Code == 10002)
                {
                    resltSell = false;
                }
                else
                {
                    Console.WriteLine($" {result.Error.Code} {result.Error.Message}");
                    Console.WriteLine();
                    if (CheckExchangeOrderStatus("ByBit", "продажа", SellSymbol, price, quantity))
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                if (resltSell == true)
                {
                    ResultTrade.Sell++;
                }
            }
            catch
            {
                Console.WriteLine();
                if (await CheckExchangeOrderStatusByBit("ByBit", "продажа", SellSymbol, price, quantity, bybitRestClient))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            return resltSell;
        }

        static async Task<bool> BuyResultMexc(MexcRestClient mexcRestClient, string BuySymbol, decimal price, decimal quantity)
        {
            bool resltBuy = true;
            try
            {
                WebCallResult<MexcOrder> result = null;
                WebCallResult<MexcOrder> resultOrderBuy = null;
                try
                {
                    result = await mexcRestClient.SpotApi.Trading.PlaceOrderAsync
                        (
                            symbol: BuySymbol,
                            side: Mexc.Net.Enums.OrderSide.Buy,
                            type: Mexc.Net.Enums.OrderType.FillOrKill,
                            quantity: quantity,
                            price: price
                        );
                }
                catch (Exception ex)
                {
                    Console.WriteLine();
                    if (await CheckExchangeOrderStatusMEXC("MEXC", "покупка", BuySymbol, price, quantity, mexcRestClient))
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }

                if (result.Error == null)
                {
                    while (true)
                    {
                        try
                        {
                            resultOrderBuy = await mexcRestClient.SpotApi.Trading.GetOrderAsync
                            (
                                symbol: BuySymbol,
                                orderId: result.Data.OrderId
                            );
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine();
                            if (await CheckExchangeOrderStatusMEXC("MEXC", "покупка", BuySymbol, price, quantity, mexcRestClient))
                            {
                                return true;
                            }
                            else
                            {
                                return false;
                            }
                        }

                        if (resultOrderBuy.Error == null)
                        {
                            if (resultOrderBuy.Data.Status == Mexc.Net.Enums.OrderStatus.Filled)
                            {
                                resltBuy = true;
                                break;
                            }
                            else if (resultOrderBuy.Data.Status == Mexc.Net.Enums.OrderStatus.Canceled)
                            {
                                resltBuy = false;
                                break;
                            }
                            await Task.Delay(2000);
                            continue;
                        }
                        else if (resultOrderBuy.Error.Code == -2013)
                        {
                            resltBuy = true;
                        }
                        else if (resultOrderBuy.Error.Message == "Request timed out")
                        {
                            if (await CheckExchangeOrderStatusMEXC("MEXC", "покупка", BuySymbol, price, quantity, mexcRestClient))
                            {
                                return true;
                            }
                            else
                            {
                                return false;
                            }
                        }
                        else
                        {
                            Console.WriteLine($" {resultOrderBuy.Error.Code} {resultOrderBuy.Error.Message}");
                            Console.WriteLine();
                            if (CheckExchangeOrderStatus("MEXC", "покупка", BuySymbol, price, quantity))
                            {
                                return true;
                            }
                            else
                            {
                                return false;
                            }
                        }
                    }

                }
                else if (result.Error.Code == 429)
                {
                    await Task.Delay(1000);
                    resltBuy = false;
                }
                else if (result.Error.Code == -2013)
                {
                    resltBuy = true;
                }
                else
                {
                    Console.WriteLine($" {result.Error.Code} {result.Error.Message}");
                    Console.WriteLine();
                    if (CheckExchangeOrderStatus("MEXC", "покупка", BuySymbol, price, quantity))
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                if (resltBuy == true)
                {
                    ResultTrade.Buy++;
                }
            }
            catch
            {
                Console.WriteLine();
                if (await CheckExchangeOrderStatusMEXC("MEXC", "покупка", BuySymbol, price, quantity, mexcRestClient))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }

            return resltBuy;
        }
        static async Task<bool> SellResultMexc(MexcRestClient mexcRestClient, string SellSymbol, decimal price, decimal quantity)
        {
            bool resltSell = true;
            try
            {
                WebCallResult<MexcOrder> result = null;
                WebCallResult<MexcOrder> resultOrderSell = null;
                try
                {
                    result = await mexcRestClient.SpotApi.Trading.PlaceOrderAsync
                        (
                            symbol: SellSymbol,
                            side: Mexc.Net.Enums.OrderSide.Sell,
                            type: Mexc.Net.Enums.OrderType.FillOrKill,
                            quantity: quantity,
                            price: price
                        );
                }
                catch (Exception ex)
                {
                    Console.WriteLine();
                    if (await CheckExchangeOrderStatusMEXC("MEXC", "продажа", SellSymbol, price, quantity, mexcRestClient))
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }

                if (result.Error == null)
                {
                    while (true)
                    {
                        try
                        {
                            resultOrderSell = await mexcRestClient.SpotApi.Trading.GetOrderAsync
                            (
                                symbol: SellSymbol,
                                orderId: result.Data.OrderId
                            );
                        }
                        catch
                        {
                            Console.WriteLine();
                            if (await CheckExchangeOrderStatusMEXC("MEXC", "продажа", SellSymbol, price, quantity, mexcRestClient))
                            {
                                return true;
                            }
                            else
                            {
                                return false;
                            }
                        }

                        if (resultOrderSell.Error == null)
                        {
                            if (resultOrderSell.Data.Status == Mexc.Net.Enums.OrderStatus.Filled)
                            {
                                resltSell = true;
                                break;
                            }
                            else if (resultOrderSell.Data.Status == Mexc.Net.Enums.OrderStatus.Canceled)
                            {
                                resltSell = false;
                                break;
                            }
                            await Task.Delay(1000); continue;
                        }
                        else if (resultOrderSell.Error.Code == -2013)
                        {
                            resltSell = true;
                        }
                        else if (resultOrderSell.Error.Message == "Request timed out")
                        {
                            if (await CheckExchangeOrderStatusMEXC("MEXC", "продажа", SellSymbol, price, quantity, mexcRestClient))
                            {
                                return true;
                            }
                            else
                            {
                                return false;
                            }
                        }
                        else
                        {
                            Console.WriteLine($" {resultOrderSell.Error.Code} {resultOrderSell.Error.Message}");
                            Console.WriteLine();
                            if (CheckExchangeOrderStatus("MEXC", "продажа", SellSymbol, price, quantity))
                            {
                                return true;
                            }
                            else
                            {
                                return false;
                            }
                        }
                    }
                }
                else if (result.Error.Code == 429)
                {
                    await Task.Delay(1000);
                    resltSell = false;
                }
                else if (result.Error.Code == -2013)
                {
                    resltSell = true;
                }
                else
                {
                    Console.WriteLine($" {result.Error.Code} {result.Error.Message}");
                    Console.WriteLine();
                    if (CheckExchangeOrderStatus("MEXC", "продажа", SellSymbol, price, quantity))
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                if (resltSell == true)
                {
                    ResultTrade.Sell++;
                }
            }
            catch
            {

            }
            return resltSell;
        }

        static public bool CheckExchangeOrderStatus(string exchange, string typeBS, string symbol, decimal price, decimal quantity)
        {
            Console.WriteLine($" Получили ошибку, зайдите на биржу: {exchange}\n" +
                              $" В историю ордеров, проверьте данный ордер:\n" +
                              $" Торговая пара: {symbol}\n" +
                              $" Цена {typeBS}: {price}\n" +
                              $" Кол-во монет: {quantity}");
            Console.WriteLine(" Исполнилась ли ваша заявка на бирже?\n" +
                              " Если исполнилась: Нажмите на кнопку (Y)\n" +
                              " Если не исполнилась: Нажмите на кнопку (N)");
            bool YesNo = false;
            while (true)
            {
                ConsoleKeyInfo keyInfo = Console.ReadKey();

                // Проверяем, какую клавишу нажал пользователь
                char keyChar = char.ToUpper(keyInfo.KeyChar); // Приводим к верхнему регистру для упрощения проверки

                if (keyChar == 'Y' || keyChar == 'Н')
                {
                    YesNo = true; break;
                }
                else if (keyChar == 'N' || keyChar == 'Т')
                {
                    YesNo = false; break;
                }
                else
                {
                    Console.WriteLine(" Вы не нажали 'Y', 'Н', 'N' или 'Т'.");
                    Console.WriteLine("Исполнилась ли ваша заявка на бирже? (Y/N):");
                }
            }
            return YesNo;
        }
        static public async Task<bool> CheckExchangeOrderStatusByBit(string exchange, string typeBS, string symbol, decimal price, decimal quantity, BybitRestClient bybitRestClient)
        {
            while (true)
            {
                try
                {
                    var history = await bybitRestClient.V5Api.Trading.GetOrderHistoryAsync(Bybit.Net.Enums.Category.Spot);
                    if (history.Error == null)
                    {
                        var order = history.Data.List.First();
                        if (order.Price == price && order.Status == Bybit.Net.Enums.V5.OrderStatus.Filled)
                        {
                            return true;
                        }
                        else
                        {
                            return false;
                        }
                    }
                    else
                    {
                        break;
                    }
                }
                catch
                {
                    await Task.Delay(1000);
                }
            }

            Console.WriteLine($" Получили ошибку, зайдите на биржу: {exchange}\n" +
                              $" В историю ордеров, проверьте данный ордер:\n" +
                              $" Торговая пара: {symbol}\n" +
                              $" Цена {typeBS}: {price}\n" +
                              $" Кол-во монет: {quantity}");
            Console.WriteLine(" Исполнилась ли ваша заявка на бирже?\n" +
                              " Если исполнилась: Нажмите на кнопку (Y)\n" +
                              " Если не исполнилась: Нажмите на кнопку (N)");
            bool YesNo = false;
            while (true)
            {
                ConsoleKeyInfo keyInfo = Console.ReadKey();

                // Проверяем, какую клавишу нажал пользователь
                char keyChar = char.ToUpper(keyInfo.KeyChar); // Приводим к верхнему регистру для упрощения проверки

                if (keyChar == 'Y' || keyChar == 'Н')
                {
                    YesNo = true; break;
                }
                else if (keyChar == 'N' || keyChar == 'Т')
                {
                    YesNo = false; break;
                }
                else
                {
                    Console.WriteLine(" Вы не нажали 'Y', 'Н', 'N' или 'Т'.");
                    Console.WriteLine("Исполнилась ли ваша заявка на бирже? (Y/N):");
                }
            }
            return YesNo;
        }
        static public async Task<bool> CheckExchangeOrderStatusMEXC(string exchange, string typeBS, string symbol, decimal price, decimal quantity, MexcRestClient mexcRestClient)
        {
            while (true)
            {
                try
                {
                    var history = await mexcRestClient.SpotApi.Trading.GetOrdersAsync(symbol);
                    if (history.Error == null)
                    {
                        var order = history.Data.First();
                        if (order.Price == price && order.Status == Mexc.Net.Enums.OrderStatus.Filled)
                        {
                            return true;
                        }
                        else
                        {
                            return false;
                        }
                    }
                    else
                    {
                        break;
                    }
                }
                catch
                {
                    await Task.Delay(1000);
                }
            }
            Console.WriteLine($" Получили ошибку, зайдите на биржу: {exchange}\n" +
                              $" В историю ордеров, проверьте данный ордер:\n" +
                              $" Торговая пара: {symbol}\n" +
                              $" Цена {typeBS}: {price}\n" +
                              $" Кол-во монет: {quantity}");
            Console.WriteLine(" Исполнилась ли ваша заявка на бирже?\n" +
                              " Если исполнилась: Нажмите на кнопку (Y)\n" +
                              " Если не исполнилась: Нажмите на кнопку (N)");
            bool YesNo = false;
            while (true)
            {
                ConsoleKeyInfo keyInfo = Console.ReadKey();

                // Проверяем, какую клавишу нажал пользователь
                char keyChar = char.ToUpper(keyInfo.KeyChar); // Приводим к верхнему регистру для упрощения проверки

                if (keyChar == 'Y' || keyChar == 'Н')
                {
                    YesNo = true; break;
                }
                else if (keyChar == 'N' || keyChar == 'Т')
                {
                    YesNo = false; break;
                }
                else
                {
                    Console.WriteLine(" Вы не нажали 'Y', 'Н', 'N' или 'Т'.");
                    Console.WriteLine("Исполнилась ли ваша заявка на бирже? (Y/N):");
                }
            }
            return YesNo;
        }
    }
}
