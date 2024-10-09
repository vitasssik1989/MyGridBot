using Bybit.Net.Clients;
using ClosedXML.Excel;
using CryptoExchange.Net.Objects;
using DocumentFormat.OpenXml.Spreadsheet;
using Mexc.Net.Clients;
using Mexc.Net.Objects.Models.Spot;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace MyGridBot
{
    internal class NewExcel
    {
        static string Path { get; set; } = @"..\\..\\..\\..\\Work";
        public static async Task TradingPairAsync(string tradingPair, BybitRestClient bybitRestClient)
        {
            string path = @"..\\..\\..\\..\\Work\\ШАБЛОН.xlsx";
            WebCallResult<Bybit.Net.Objects.Models.V5.BybitResponse<Bybit.Net.Objects.Models.V5.BybitSpotSymbol>> symbolData;
            while (true)
            {
                try
                {
                    symbolData = await bybitRestClient.V5Api.ExchangeData.GetSpotSymbolsAsync();

                    if (symbolData.ResponseStatusCode == System.Net.HttpStatusCode.OK)
                    {
                        break;
                    }
                }
                catch
                {

                }
            }
            foreach (var symbol in symbolData.Data.List)
            {
                if (symbol.Name == tradingPair)
                {
                    while (true)
                    {
                        try
                        {
                            using (var workbook = new XLWorkbook(path))
                            {
                                var sheet = workbook.Worksheet(1);
                                string formatCommaPrice = FormatZeroСomma(symbol.PriceFilter.TickSize);
                                string formatCommaBase = FormatZeroСomma(symbol.LotSizeFilter.BasePrecision);
                                sheet.Cell(2, 15).Value = ValueAfterComma(symbol.LotSizeFilter.BasePrecision);
                                //decimal minQty = Math.Round(symbol.MinOrderQuantity + (symbol.MinOrderQuantity / 100 * 0.1m), formatCommaBase.Length - 2);
                                decimal minQty = symbol.LotSizeFilter.MinOrderQuantity;
                                sheet.Cell(14, 16).Value = symbol.LotSizeFilter.MinOrderQuantity;
                                sheet.Cell(2, 8).Value = minQty;
                                while (Convert.ToDecimal(sheet.Cell(2, 7).Value) < symbol.LotSizeFilter.MinOrderQuantity)
                                {
                                    minQty += symbol.LotSizeFilter.BasePrecision;
                                    sheet.Cell(2, 8).Value = minQty;
                                }
                                for (int i = 2; i <= 5001; i++)
                                {
                                    sheet.Cell(i, 1).Value = 0;
                                    sheet.Cell(i, 4).Value = 0;
                                    sheet.Cell(i, 5).Value = 0;
                                    sheet.Cell(i, 6).Value = 0;


                                    sheet.Cell(i, 8).AddConditionalFormat().WhenIsTrue($"=D{i}=1")
                                    .Fill.SetBackgroundColor(XLColor.FromHtml("#ffa770"));
                                    sheet.Cell(i, 8).AddConditionalFormat().WhenIsTrue($"=A{i}=1")
                                   .Fill.SetBackgroundColor(XLColor.FromHtml("#a8ffc5"));
                                    sheet.Cell(i, 8).AddConditionalFormat().WhenIsTrue($"=Q{i}=1")
                                   .Fill.SetBackgroundColor(XLColor.FromHtml("#FF2400"));

                                    sheet.Cell(i, 9).Style.NumberFormat.Format = "0.00000000000000000000";
                                    sheet.Cell(i, 10).Style.NumberFormat.Format = "0.00000000000000000000";
                                    sheet.Cell(i, 14).Style.NumberFormat.Format = "0.000000000000000000000000000000";

                                    sheet.Cell(i, 2).Style.NumberFormat.Format = formatCommaPrice;//PricePrecision
                                    sheet.Cell(i, 3).Style.NumberFormat.Format = formatCommaPrice;//PricePrecision
                                    sheet.Cell(i, 2).Value = 0;

                                    sheet.Cell(i, 7).Style.NumberFormat.Format = formatCommaBase; //BasePrecision
                                    sheet.Cell(i, 8).Style.NumberFormat.Format = formatCommaBase; //BasePrecision

                                    sheet.Cell(i, 8).Value = minQty;
                                }
                                sheet.Cell(2, 16).Value = CountTrailingZerosAfterDecimal(symbol.PriceFilter.TickSize.ToString());
                                workbook.SaveAs($@"..\\..\\..\\..\\Work\\{tradingPair}.xlsx");
                            }
                            break;
                        }
                        catch
                        {
                            Thread.Sleep(10000);
                            Console.WriteLine(" Не смог открыть ексель ШАБЛОН в папке Work");
                        }
                    }
                    break;
                }
            }
        }
        public static async Task TradingPairAsyncMexc(string tradingPair, MexcRestClient mexcRestClient)
        {
            string path = @"..\\..\\..\\..\\WorkMexc\\ШАБЛОН.xlsx";
            WebCallResult<MexcExchangeInfo> symbolData;
            while (true)
            {
                try
                {
                    symbolData = await mexcRestClient.SpotApi.ExchangeData.GetExchangeInfoAsync();

                    if (symbolData.ResponseStatusCode == System.Net.HttpStatusCode.OK)
                    {
                        break;
                    }
                }
                catch
                {

                }
            }

            while (true)
            {
                foreach (var symbol in symbolData.Data.Symbols)
                {
                    if (symbol.Name == tradingPair)
                    {
                        if (!symbol.IsSpotTradingAllowed)
                        {
                            Console.WriteLine($" Пара: {tradingPair} не торгуется через API\n" +
                                              $" Нажмите ENTER");
                            Console.ReadLine();
                            return;
                        }
                        else { break; }
                    }
                }
                break;
            }

            foreach (var symbol in symbolData.Data.Symbols)
            {
                if (symbol.Name == tradingPair)
                {
                    while (true)
                    {
                        try
                        {
                            using (var workbook = new XLWorkbook(path))
                            {
                                var sheet = workbook.Worksheet(1);

                                string formatCommaPrice = FormatZeroСommaMexc(symbol.QuoteAssetPrecision);
                                string formatCommaBase = FormatZeroСommaMexc(symbol.BaseAssetPrecision);
                                sheet.Cell(2, 15).Value = symbol.BaseAssetPrecision;
                                sheet.Cell(14, 16).Value = 0;
                                for (int i = 2; i <= 5001; i++)
                                {
                                    sheet.Cell(i, 1).Value = 0;
                                    sheet.Cell(i, 4).Value = 0;
                                    sheet.Cell(i, 5).Value = 0;
                                    sheet.Cell(i, 6).Value = 0;


                                    sheet.Cell(i, 8).AddConditionalFormat().WhenIsTrue($"=D{i}=1")
                                    .Fill.SetBackgroundColor(XLColor.FromHtml("#ffa770"));
                                    sheet.Cell(i, 8).AddConditionalFormat().WhenIsTrue($"=A{i}=1")
                                   .Fill.SetBackgroundColor(XLColor.FromHtml("#a8ffc5"));
                                    sheet.Cell(i, 8).AddConditionalFormat().WhenIsTrue($"=Q{i}=1")
                                   .Fill.SetBackgroundColor(XLColor.FromHtml("#FF2400"));

                                    sheet.Cell(i, 9).Style.NumberFormat.Format = "0.00000000000000000000";
                                    sheet.Cell(i, 10).Style.NumberFormat.Format = "0.00000000000000000000";
                                    sheet.Cell(i, 14).Style.NumberFormat.Format = "0.000000000000000000000000000000";

                                    sheet.Cell(i, 2).Style.NumberFormat.Format = formatCommaPrice;//PricePrecision
                                    sheet.Cell(i, 3).Style.NumberFormat.Format = formatCommaPrice;//PricePrecision
                                    sheet.Cell(i, 2).Value = 0;

                                    sheet.Cell(i, 7).Style.NumberFormat.Format = formatCommaBase; //BasePrecision
                                    sheet.Cell(i, 8).Style.NumberFormat.Format = formatCommaBase; //BasePrecision

                                    sheet.Cell(i, 8).Value = 0;
                                }
                                sheet.Cell(2, 16).Value = symbol.QuoteAssetPrecision;
                                workbook.SaveAs($@"..\\..\\..\\..\\WorkMexc\\{tradingPair}.xlsx");
                            }
                            break;
                        }
                        catch
                        {
                            Thread.Sleep(10000);
                            Console.WriteLine(" Не смог открыть ексель ШАБЛОН в папке WorkMexc");
                        }
                    }
                    break;
                }
            }
        }
        public static async Task Setka(string tradingPair, BybitRestClient bybitRestClient)
        {
            WebCallResult<Bybit.Net.Objects.Models.V5.BybitResponse<Bybit.Net.Objects.Models.V5.BybitSpotSymbol>> symbolData;
            while (true)
            {
                try
                {
                    symbolData = await bybitRestClient.V5Api.ExchangeData.GetSpotSymbolsAsync();

                    if (symbolData.ResponseStatusCode == System.Net.HttpStatusCode.OK)
                    {
                        break;
                    }
                }
                catch
                {

                }
            }
            foreach (var symbol in symbolData.Data.List)
            {
                if (symbol.Name == tradingPair)
                {
                    while (true)
                    {
                        try
                        {

                            using (var workbook = new XLWorkbook(@$"..\\..\\..\\..\\Work\\{tradingPair}.xlsx"))
                            {
                                var sheet = workbook.Worksheet(1);
                                Console.WriteLine();
                                Console.WriteLine($" Введите максимальную цену и нажмите ENTER\n" +
                                                  $" Пример ввода: {symbol.PriceFilter.TickSize} ");
                                decimal haigPrice = Kultura(Console.ReadLine());

                                Console.WriteLine($" Введите шаг цены и нажмите ENTER\n" +
                                                  $" Пример ввода: {symbol.PriceFilter.TickSize}");
                                decimal priceStep = Kultura(Console.ReadLine());
                                while (true)
                                {
                                    if (priceStep < symbol.PriceFilter.TickSize || CountTrailingZerosAfterDecimal(priceStep.ToString()) > CountTrailingZerosAfterDecimal(symbol.PriceFilter.TickSize.ToString()))
                                    {
                                        Console.WriteLine($" Неверно указано кол-во символов: \n" +
                                                          $" Пример: {symbol.PriceFilter.TickSize} \n" +
                                                          $" Вы ввели: {priceStep}\n" +
                                                          $" Введите шаг цены и нажмите ENTER");
                                        priceStep = Kultura(Console.ReadLine());
                                    }
                                    else
                                    {
                                        break;
                                    }
                                }
                                Console.WriteLine($" Введите процент продажи от цены покупки(спред) и нажмите ENTER\n" +
                                                  $" Пример ввода: 2,125");
                                decimal precent = Kultura(Console.ReadLine());
                                sheet.Cell(2, 16).Value = CountTrailingZerosAfterDecimal(symbol.PriceFilter.TickSize.ToString());
                                sheet.Cell(8, 16).Value = precent;

                                for (int i = 2; i <= 5001; i++)
                                {
                                    if (haigPrice > 0)
                                    {
                                        sheet.Cell(i, 2).Value = haigPrice;
                                        haigPrice -= priceStep;
                                    }
                                    else { break; }
                                }
                                workbook.Save();
                                Console.WriteLine(" ВНИМАТЕЛЬНО ПРОЧИТАЙ И НАЖМИ ПОТОМ ENTER\n" +
                                                  " Спред(%) продажи можешь устанавливать\n" +
                                                  " в ячейке P8");
                                Console.ReadLine();
                            }
                            break;
                        }
                        catch
                        {
                            Thread.Sleep(10000);
                            Console.WriteLine($" Не смог открыть ексель {tradingPair}.xlsx в папке Work");
                        }
                    }
                    break;
                }
            }
        }
        public static async Task SetkaMexc(string tradingPair, MexcRestClient mexcRestClient)
        {
            WebCallResult<MexcExchangeInfo> symbolData;
            while (true)
            {
                try
                {
                    symbolData = await mexcRestClient.SpotApi.ExchangeData.GetExchangeInfoAsync();

                    if (symbolData.ResponseStatusCode == System.Net.HttpStatusCode.OK)
                    {
                        break;
                    }
                }
                catch
                {

                }
            }
            foreach (var symbol in symbolData.Data.Symbols)
            {
                if (symbol.Name == tradingPair)
                {
                    while (true)
                    {
                        try
                        {

                            using (var workbook = new XLWorkbook(@$"..\\..\\..\\..\\WorkMexc\\{tradingPair}.xlsx"))
                            {
                                var sheet = workbook.Worksheet(1);
                                Console.WriteLine();
                                Console.WriteLine($" Введите максимальную цену и нажмите ENTER\n" +
                                                  $" Не превышая знаков после запятой: {symbol.QuoteAssetPrecision}");
                                decimal haigPrice = Kultura(Console.ReadLine());

                                Console.WriteLine($" Введите шаг цены и нажмите ENTER\n" +
                                                  $" Не превышая знаков после запятой: {symbol.QuoteAssetPrecision}");
                                decimal priceStep = Kultura(Console.ReadLine());

                                sheet.Cell(2, 16).Value = symbol.QuoteAssetPrecision;

                                for (int i = 2; i <= 5001; i++)
                                {
                                    if (haigPrice > 0)
                                    {
                                        sheet.Cell(i, 2).Value = haigPrice;
                                        haigPrice -= priceStep;
                                    }
                                    else { break; }
                                }
                                workbook.Save();
                            }
                            break;
                        }
                        catch
                        {
                            Thread.Sleep(10000);
                            Console.WriteLine($" Не смог открыть ексель {tradingPair}.xlsx в папке WorkMexc");
                        }
                    }
                    break;
                }
            }
        }
        public static async Task SortBuySellByBitAsync(BybitRestClient bybitRestClient)
        {
            List<decimal> sortBS = new List<decimal>();
            foreach (var excelSort in SettingStart.SymbolList)
            {
                await DynamicSort(bybitRestClient, excelSort);
                try
                {
                    using (var workbook = new XLWorkbook(@$"..\\..\\..\\..\\Work\\{excelSort}.xlsx"))
                    {
                        var sheet = workbook.Worksheet(1);

                        //Обычная сортировка
                        if (!sheet.Cell(7, 16).IsEmpty())
                        {
                            int sort = Convert.ToInt32(sheet.Cell(7, 16).Value);
                            if (sort > 0)
                            {
                                int playsort = 0;
                                for (int i = 2; i <= 5001; i++)
                                {
                                    if (Convert.ToInt32(sheet.Cell(i, 17).Value) == 1)
                                    {
                                        playsort++;
                                    }
                                }
                                if (playsort == 2)
                                {
                                    if (sort > 0)
                                    {
                                        Console.WriteLine();
                                        Console.Write(" Сортировка.Пара: ");
                                        Console.ForegroundColor = ConsoleColor.White;
                                        Console.Write($"{excelSort}");
                                        Console.ForegroundColor = ConsoleColor.Blue;
                                        Console.WriteLine();

                                        //Сортировка ордеров на продажу
                                        if (sort == 1 || sort == 2)
                                        {
                                            Console.Write(" Сортирую ордера! Тип: ");
                                            Console.ForegroundColor = ConsoleColor.Red;
                                            Console.Write($"Sell");
                                            Console.ForegroundColor = ConsoleColor.Blue;
                                            Console.WriteLine();
                                            for (int s = 0; s < 2; s++)
                                            {
                                                if (s == 0)
                                                {
                                                    int flag = 0;
                                                    for (int i = 2; i <= 5001; i++)
                                                    {
                                                        if (Convert.ToInt32(sheet.Cell(i, 17).Value) == 1)
                                                        {
                                                            if (flag == 0)
                                                            {
                                                                flag = 1;
                                                            }
                                                            else { flag = 2; }
                                                        }

                                                        if (flag > 0)
                                                        {
                                                            if (Convert.ToInt32(sheet.Cell(i, 4).Value) == 1 && Convert.ToInt32(sheet.Cell(i, 1).Value) == 1)
                                                            {
                                                                sortBS.Add(Convert.ToDecimal(sheet.Cell(i, 8).Value));
                                                            }
                                                        }
                                                        if (flag == 2) { break; }
                                                    }
                                                    if (sortBS.Count > 0)
                                                    {
                                                        sortBS.Sort((a, b) => b.CompareTo(a));
                                                    }
                                                    else
                                                    {
                                                        Console.WriteLine($" Нет ордеров! Тип: Sell");
                                                        break;
                                                    }
                                                }
                                                else
                                                {
                                                    int sortIndex = 0;
                                                    int flag = 0;
                                                    for (int i = 2; i <= 5001; i++)
                                                    {

                                                        if (Convert.ToInt32(sheet.Cell(i, 17).Value) == 1)
                                                        {
                                                            if (flag == 0)
                                                            {
                                                                flag = 1;
                                                            }
                                                            else { flag = 2; }
                                                        }

                                                        if (flag > 0)
                                                        {
                                                            if (Convert.ToInt32(sheet.Cell(i, 4).Value) == 1 && Convert.ToInt32(sheet.Cell(i, 1).Value) == 1)
                                                            {
                                                                sheet.Cell(i, 8).Value = sortBS[sortIndex];
                                                                sortIndex++;
                                                            }
                                                        }
                                                        if (flag == 2) { break; }
                                                    }
                                                    workbook.Save();
                                                    Console.Write(" Кол-во ордеров: ");
                                                    Console.ForegroundColor = ConsoleColor.Red;
                                                    Console.Write($"{sortBS.Count}");
                                                    Console.ForegroundColor = ConsoleColor.Blue;
                                                    Console.WriteLine();
                                                    sortBS.Clear();
                                                }
                                            }
                                        }

                                        //Сортировка ордеров на покупку
                                        if (sort == 1 || sort == 3)
                                        {
                                            Console.Write(" Сортирую ордера! Тип: ");
                                            Console.ForegroundColor = ConsoleColor.Green;
                                            Console.Write($"Buy");
                                            Console.ForegroundColor = ConsoleColor.Blue;
                                            Console.WriteLine();
                                            for (int s = 0; s < 2; s++)
                                            {
                                                if (s == 0)
                                                {
                                                    int flag = 0;
                                                    for (int i = 5001; i >= 2; i--)
                                                    {
                                                        if (Convert.ToInt32(sheet.Cell(i, 17).Value) == 1)
                                                        {
                                                            if (flag == 0)
                                                            {
                                                                flag = 1;
                                                            }
                                                            else { flag = 2; }
                                                        }

                                                        if (flag > 0)
                                                        {
                                                            if (Convert.ToInt32(sheet.Cell(i, 4).Value) == 0 && Convert.ToInt32(sheet.Cell(i, 1).Value) == 1)
                                                            {
                                                                sortBS.Add(Convert.ToDecimal(sheet.Cell(i, 8).Value));
                                                            }
                                                        }
                                                        if (flag == 2) { break; }
                                                    }
                                                    if (sortBS.Count > 0)
                                                    {
                                                        sortBS.Sort((a, b) => b.CompareTo(a));
                                                    }
                                                    else
                                                    {
                                                        Console.WriteLine($" Нет ордеров! Тип: Buy");
                                                        break;
                                                    }
                                                }
                                                else
                                                {
                                                    int sortIndex = 0;
                                                    int flag = 0;
                                                    for (int i = 5001; i >= 2; i--)
                                                    {
                                                        if (Convert.ToInt32(sheet.Cell(i, 17).Value) == 1)
                                                        {
                                                            if (flag == 0)
                                                            {
                                                                flag = 1;
                                                            }
                                                            else { flag = 2; }
                                                        }

                                                        if (flag > 0)
                                                        {
                                                            if (Convert.ToInt32(sheet.Cell(i, 4).Value) == 0 && Convert.ToInt32(sheet.Cell(i, 1).Value) == 1)
                                                            {
                                                                sheet.Cell(i, 8).Value = sortBS[sortIndex];
                                                                sortIndex++;
                                                            }
                                                        }
                                                        if (flag == 2) { break; }
                                                    }
                                                    workbook.Save();
                                                    Console.Write(" Кол-во ордеров: ");
                                                    Console.ForegroundColor = ConsoleColor.Green;
                                                    Console.Write($"{sortBS.Count}");
                                                    Console.ForegroundColor = ConsoleColor.Blue;
                                                    Console.WriteLine();
                                                    sortBS.Clear();
                                                }
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    Console.WriteLine($" Неверно указан диапазон в паре {excelSort}");
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(" Не смог открыть ексель");
                    Console.WriteLine(ex.Message);
                    Console.ReadLine();
                }
            }
        }
        public static async Task SortBuySellMexcAsync()
        {
            List<decimal> sortBS = new List<decimal>();
            foreach (var excelSort in SettingStart.SymbolList)
            {
                try
                {
                    using (var workbook = new XLWorkbook(@$"..\\..\\..\\..\\WorkMexc\\{excelSort}.xlsx"))
                    {
                        var sheet = workbook.Worksheet(1);

                        //Обычная сортировка
                        if (!sheet.Cell(7, 16).IsEmpty())
                        {
                            int sort = Convert.ToInt32(sheet.Cell(7, 16).Value);
                            if (sort > 0)
                            {
                                int playsort = 0;
                                for (int i = 2; i <= 5001; i++)
                                {
                                    if (Convert.ToInt32(sheet.Cell(i, 17).Value) == 1)
                                    {
                                        playsort++;
                                    }
                                }
                                if (playsort == 2)
                                {
                                    if (sort > 0)
                                    {
                                        Console.WriteLine();
                                        Console.Write(" Сортировка.Пара: ");
                                        Console.ForegroundColor = ConsoleColor.White;
                                        Console.Write($"{excelSort}");
                                        Console.ForegroundColor = ConsoleColor.Blue;
                                        Console.WriteLine();

                                        //Сортировка ордеров на продажу
                                        if (sort == 1 || sort == 2)
                                        {
                                            Console.Write(" Сортирую ордера! Тип: ");
                                            Console.ForegroundColor = ConsoleColor.Red;
                                            Console.Write($"Sell");
                                            Console.ForegroundColor = ConsoleColor.Blue;
                                            Console.WriteLine();
                                            for (int s = 0; s < 2; s++)
                                            {
                                                if (s == 0)
                                                {
                                                    int flag = 0;
                                                    for (int i = 2; i <= 5001; i++)
                                                    {
                                                        if (Convert.ToInt32(sheet.Cell(i, 17).Value) == 1)
                                                        {
                                                            if (flag == 0)
                                                            {
                                                                flag = 1;
                                                            }
                                                            else { flag = 2; }
                                                        }

                                                        if (flag > 0)
                                                        {
                                                            if (Convert.ToInt32(sheet.Cell(i, 4).Value) == 1 && Convert.ToInt32(sheet.Cell(i, 1).Value) == 1)
                                                            {
                                                                sortBS.Add(Convert.ToDecimal(sheet.Cell(i, 8).Value));
                                                            }
                                                        }
                                                        if (flag == 2) { break; }
                                                    }
                                                    if (sortBS.Count > 0)
                                                    {
                                                        sortBS.Sort((a, b) => b.CompareTo(a));
                                                    }
                                                    else
                                                    {
                                                        Console.WriteLine($" Нет ордеров! Тип: Sell");
                                                        break;
                                                    }
                                                }
                                                else
                                                {
                                                    int sortIndex = 0;
                                                    int flag = 0;
                                                    for (int i = 2; i <= 5001; i++)
                                                    {

                                                        if (Convert.ToInt32(sheet.Cell(i, 17).Value) == 1)
                                                        {
                                                            if (flag == 0)
                                                            {
                                                                flag = 1;
                                                            }
                                                            else { flag = 2; }
                                                        }

                                                        if (flag > 0)
                                                        {
                                                            if (Convert.ToInt32(sheet.Cell(i, 4).Value) == 1 && Convert.ToInt32(sheet.Cell(i, 1).Value) == 1)
                                                            {
                                                                sheet.Cell(i, 8).Value = sortBS[sortIndex];
                                                                sortIndex++;
                                                            }
                                                        }
                                                        if (flag == 2) { break; }
                                                    }
                                                    workbook.Save();
                                                    Console.Write(" Кол-во ордеров: ");
                                                    Console.ForegroundColor = ConsoleColor.Red;
                                                    Console.Write($"{sortBS.Count}");
                                                    Console.ForegroundColor = ConsoleColor.Blue;
                                                    Console.WriteLine();
                                                    sortBS.Clear();
                                                }
                                            }
                                        }

                                        //Сортировка ордеров на покупку
                                        if (sort == 1 || sort == 3)
                                        {
                                            Console.Write(" Сортирую ордера! Тип: ");
                                            Console.ForegroundColor = ConsoleColor.Green;
                                            Console.Write($"Buy");
                                            Console.ForegroundColor = ConsoleColor.Blue;
                                            Console.WriteLine();
                                            for (int s = 0; s < 2; s++)
                                            {
                                                if (s == 0)
                                                {
                                                    int flag = 0;
                                                    for (int i = 5001; i >= 2; i--)
                                                    {
                                                        if (Convert.ToInt32(sheet.Cell(i, 17).Value) == 1)
                                                        {
                                                            if (flag == 0)
                                                            {
                                                                flag = 1;
                                                            }
                                                            else { flag = 2; }
                                                        }

                                                        if (flag > 0)
                                                        {
                                                            if (Convert.ToInt32(sheet.Cell(i, 4).Value) == 0 && Convert.ToInt32(sheet.Cell(i, 1).Value) == 1)
                                                            {
                                                                sortBS.Add(Convert.ToDecimal(sheet.Cell(i, 8).Value));
                                                            }
                                                        }
                                                        if (flag == 2) { break; }
                                                    }
                                                    if (sortBS.Count > 0)
                                                    {
                                                        sortBS.Sort((a, b) => b.CompareTo(a));
                                                    }
                                                    else
                                                    {
                                                        Console.WriteLine($" Нет ордеров! Тип: Buy");
                                                        break;
                                                    }
                                                }
                                                else
                                                {
                                                    int sortIndex = 0;
                                                    int flag = 0;
                                                    for (int i = 5001; i >= 2; i--)
                                                    {
                                                        if (Convert.ToInt32(sheet.Cell(i, 17).Value) == 1)
                                                        {
                                                            if (flag == 0)
                                                            {
                                                                flag = 1;
                                                            }
                                                            else { flag = 2; }
                                                        }

                                                        if (flag > 0)
                                                        {
                                                            if (Convert.ToInt32(sheet.Cell(i, 4).Value) == 0 && Convert.ToInt32(sheet.Cell(i, 1).Value) == 1)
                                                            {
                                                                sheet.Cell(i, 8).Value = sortBS[sortIndex];
                                                                sortIndex++;
                                                            }
                                                        }
                                                        if (flag == 2) { break; }
                                                    }
                                                    workbook.Save();
                                                    Console.Write(" Кол-во ордеров: ");
                                                    Console.ForegroundColor = ConsoleColor.Green;
                                                    Console.Write($"{sortBS.Count}");
                                                    Console.ForegroundColor = ConsoleColor.Blue;
                                                    Console.WriteLine();
                                                    sortBS.Clear();
                                                }
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    Console.WriteLine($" Неверно указан диапазон в паре {excelSort}");
                                }
                            }

                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(" Не смог открыть ексель");
                    Console.WriteLine(ex.Message);
                    Console.ReadLine();
                }
            }
        }
        public static async Task DynamicSort(BybitRestClient bybitRestClient, string excelSort)
        {
            try
            {
                using (var workbook = new XLWorkbook(@$"..\\..\\..\\..\\Work\\{excelSort}.xlsx"))
                {
                    var sheet = workbook.Worksheet(1);
                    //Динамическая сортировка и HL

                    if (!sheet.Cell(10, 15).IsEmpty() && !sheet.Cell(10, 16).IsEmpty() && !sheet.Cell(12, 15).IsEmpty() && !sheet.Cell(12, 16).IsEmpty() && !sheet.Cell(7, 16).IsEmpty())
                    {
                        if (Convert.ToInt32(sheet.Cell(7, 16).Value) > 0)
                        {
                            int DynamicSortBUY = Convert.ToInt32(sheet.Cell(10, 15).Value);
                            int DynamicSortSELL = Convert.ToInt32(sheet.Cell(10, 16).Value);
                            int DynamicHigh = Convert.ToInt32(sheet.Cell(12, 15).Value);
                            int DynamicLow = Convert.ToInt32(sheet.Cell(12, 16).Value);

                            if (DynamicSortBUY > 0 && DynamicSortSELL > 0)
                            {
                                Console.WriteLine();
                                Console.ForegroundColor = ConsoleColor.Blue;
                                Console.Write($" Динамическая Сортировка: ");
                                Console.ForegroundColor = ConsoleColor.White;
                                Console.Write($"{excelSort}\n");
                                Console.ForegroundColor = ConsoleColor.Blue;
                                if (DynamicSortBUY + DynamicSortSELL > DynamicHigh + DynamicLow)
                                {

                                    decimal PriceReal = (await Trader.AskPriceQuantityByBit(bybitRestClient, excelSort)).Price;

                                    Console.Write($" DynamicSort Sell: ");
                                    Console.ForegroundColor = ConsoleColor.Green;
                                    Console.Write($"{DynamicSortSELL}\n");
                                    Console.ForegroundColor = ConsoleColor.Blue;
                                    Console.Write($" DynamicSort Buy: ");
                                    Console.ForegroundColor = ConsoleColor.Red;
                                    Console.Write($"{DynamicSortBUY}\n");
                                    Console.ForegroundColor = ConsoleColor.Blue;
                                    Console.Write($" Dynamic High: ");
                                    Console.ForegroundColor = ConsoleColor.Red;
                                    Console.Write($"{DynamicHigh}\n");
                                    Console.ForegroundColor = ConsoleColor.Blue;
                                    Console.Write($" Dynamic Low: ");
                                    Console.ForegroundColor = ConsoleColor.Green;
                                    Console.Write($"{DynamicLow}");
                                    Console.ForegroundColor = ConsoleColor.Blue;


                                    int str = 0;
                                    if (DynamicSortBUY > 0 && DynamicSortSELL > 0)
                                    {
                                        for (int i = 2; i <= 5001; i++)
                                        {
                                            sheet.Cell(i, 17).Value = 0;
                                            if (str == 0 && Convert.ToDecimal(sheet.Cell(i, 2).Value) < PriceReal)
                                            {
                                                str = i;
                                            }
                                        }

                                        if (str + DynamicSortBUY <= 5001)
                                        {
                                            sheet.Cell(str + DynamicSortBUY, 17).Value = 1;
                                        }
                                        else
                                        {
                                            sheet.Cell(5001, 17).Value = 1;
                                        }

                                        if (str - DynamicSortSELL >= 2)
                                        {
                                            sheet.Cell(str - DynamicSortSELL, 17).Value = 1;
                                        }
                                        else
                                        {
                                            sheet.Cell(2, 17).Value = 1;
                                        }
                                        workbook.Save();
                                    }

                                    if (DynamicHigh > 0 && DynamicLow > 0)
                                    {
                                        for (int i = 2; i <= 5001; i++)
                                        {
                                            sheet.Cell(i, 6).Value = 0;
                                        }

                                        if (str + DynamicSortBUY <= 5001)
                                        {
                                            for (int i = 0; i < DynamicLow; i++)
                                            {
                                                sheet.Cell(str + DynamicSortBUY - i, 6).Value = 2;
                                            }
                                        }
                                        else
                                        {
                                            if (DynamicLow - (str + DynamicSortBUY - 5001) > 0)
                                            {
                                                for (int i = 0; i < DynamicLow - (str + DynamicSortBUY - 5001); i++)
                                                {
                                                    sheet.Cell(5001 - i, 6).Value = 2;
                                                }
                                            }
                                        }

                                        if (str - DynamicSortSELL >= 2) // str = 100  DynamicSortSELL = 150  DynamicHigh = 85
                                        {
                                            for (int i = 0; i < DynamicHigh; i++)
                                            {
                                                sheet.Cell(str - DynamicSortSELL + i, 6).Value = 1;
                                            }
                                        }
                                        else
                                        {
                                            if (DynamicHigh - (str - DynamicSortSELL) > 0)
                                            {
                                                for (int i = 0; i < DynamicHigh - (str - DynamicSortSELL); i++)
                                                {
                                                    sheet.Cell(2 + i, 6).Value = 1;
                                                }
                                            }
                                        }
                                        workbook.Save();
                                    }

                                    Console.WriteLine();
                                    Console.ForegroundColor = ConsoleColor.Blue;
                                    Console.Write($" Динамическая Сортировка: ");
                                    Console.ForegroundColor = ConsoleColor.White;
                                    Console.Write($"{excelSort} Успех\n");
                                    Console.ForegroundColor = ConsoleColor.Blue;
                                    Console.WriteLine();
                                }
                                else
                                {
                                    Console.ForegroundColor = ConsoleColor.Red;
                                    Console.WriteLine($" Строки \n" +
                                                      $" DynamicSort BUY + DynamicSort Sell = {DynamicSortSELL + DynamicSortBUY} \n" +
                                                      $" Должны быть больше чем, чем строки\n" +
                                                      $" Dynamic High + Dynamic Low = {DynamicHigh + DynamicLow}\n" +
                                                      $" Динамическая сортировка не отработала");
                                    Console.ForegroundColor = ConsoleColor.Blue;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(" Не смог открыть ексель");
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
        }
        static string FormatZeroСomma(decimal PricePrecision)
        {
            string priceP = PricePrecision.ToString();
            string result = "";
            foreach (var item in priceP)
            {

                if (item == '1')
                {
                    result += '0';
                    continue;
                }
                if (item == ',')
                {
                    result += '.';
                    continue;
                }
                result += '0';
            }
            return result;
        }
        static string FormatZeroСommaMexc(decimal PricePrecision)
        {
            string result = "0";
            if (PricePrecision > 0)
            {
                result += '.';
                for (int i = 0; i < PricePrecision; i++)
                {
                    result += '0';
                }
            }
            return result;
        }
        static int ValueAfterComma(decimal BasePrecision)
        {
            string a = BasePrecision.ToString();
            int result = 0;
            if (a.Length == 1)
            {
                return 0;
            }
            return a.Length - 2;
        }
        static decimal Kultura(string kultyra)
        {
            decimal result = 0;
            if (decimal.TryParse(kultyra.Replace(',', '.'), out decimal H))
            {
                result = H;
            }
            else
            {
                if (decimal.TryParse(kultyra.Replace('.', ','), out decimal Hh))
                {
                    result = Hh;
                }
            }
            return result;
        }
        public static int CountTrailingZerosAfterDecimal(string input)
        {
            // Находим индекс разделителя (запятой или точки)
            int decimalIndex = input.IndexOfAny(new char[] { ',', '.' });

            // Если запятая не найдена или она в конце строки, возвращаем 0
            if (decimalIndex == -1 || decimalIndex == input.Length - 1)
                return 0;

            // Считаем количество нулей после запятой
            int zeroCount = 0;
            for (int i = input.Length - 1; i > decimalIndex; i--)
            {
                if (input[i] == '.' && input[i] == ',')
                {
                    break;
                }
                zeroCount++;
            }

            return zeroCount;
        }
    }
}
