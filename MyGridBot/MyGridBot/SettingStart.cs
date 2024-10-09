using Bybit.Net.Clients;
using ClosedXML.Excel;
using CryptoExchange.Net.Objects;
using Mexc.Net.Clients;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyGridBot
{
    internal class SettingStart
    {
        static string _path = @"..\\..\\..\\..\\Work\\Setting.xlsx";
        static string _patnMexc = @"..\\..\\..\\..\\WorkMexc\\Setting.xlsx";
        public static string APIkey { get; set; }
        public static string APIsecret { get; set; }

        public static List<string> SymbolList;

        public static void Start()
        {
            Console.WriteLine(" Открываю ексель Setting.xlsx в папке Work");

            while (true)
            {
                try
                {
                    while (true)
                    {
                        using (var workbook = new XLWorkbook(_path))
                        {
                            var sheet = workbook.Worksheet(1);

                            if (sheet.Cell(1, 3).IsEmpty() && sheet.Cell(2, 3).IsEmpty())
                            {
                                Console.WriteLine(" Укажите APIkey и APIsecret и нажмите ENTER");
                                Console.ReadLine();
                                workbook.Dispose();
                                continue;
                            }
                            APIkey = sheet.Cell(1, 3).Value.ToString();
                            APIsecret = sheet.Cell(2, 3).Value.ToString();
                            break;
                        }
                    }
                    break;
                }
                catch
                {
                    Console.WriteLine(" Не смог открыть ексель Setting.xlsx в папке Work\n" +
                                      " Проверь не открыта ли ексель или есть ли доступ");
                    Thread.Sleep(10000);
                }
            }
        }
        public static void StartMexc()
        {
            Console.WriteLine(" Открываю ексель Setting.xlsx в папке WorkMexc");

            while (true)
            {
                try
                {
                    while (true)
                    {
                        using (var workbook = new XLWorkbook(_patnMexc))
                        {
                            var sheet = workbook.Worksheet(1);
                            if (sheet.Cell(1, 3).IsEmpty() && sheet.Cell(2, 3).IsEmpty())
                            {
                                Console.WriteLine(" Укажите APIkey и APIsecret и нажмите ENTER");
                                Console.ReadLine();
                                workbook.Dispose();
                                continue;
                            }
                            APIkey = sheet.Cell(1, 3).Value.ToString();
                            APIsecret = sheet.Cell(2, 3).Value.ToString();
                            break;
                        }
                    }
                    break;
                }
                catch
                {
                    Console.WriteLine(" Не смог открыть ексель Setting.xlsx в папке WorkMexc\n" +
                                      " Проверь не открыта ли ексель или есть ли доступ");
                    Thread.Sleep(10000);
                }
            }
        }
        public static void UpdateSymbolList()
        {
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.WriteLine();
            Console.WriteLine(" Копирую все торговые пары");
            while (true)
            {
                try
                {
                    SymbolList = new List<string>();
                    using (var workbook = new XLWorkbook(_path))
                    {
                        var sheet = workbook.Worksheet(1);
                        for (int i = 2; i < 500; i++)
                        {
                            if (sheet.Cell(i, 1).IsEmpty() != true)
                            {
                                SymbolList.Add(sheet.Cell(i, 1).Value.ToString());
                            }
                            else { break; }
                        }
                    }
                    break;
                }
                catch
                {
                    Console.WriteLine(" Не смог открыть ексель Setting.xlsx в папке Work\n" +
                                      " Проверь не открыта ли ексель или есть ли доступ");
                    Thread.Sleep(10000);
                }
            }
        }
        public static void UpdateSymbolListMexc()
        {
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.WriteLine();
            Console.WriteLine(" Копирую все торговые пары");
            while (true)
            {
                try
                {
                    SymbolList = new List<string>();
                    using (var workbook = new XLWorkbook(_patnMexc))
                    {
                        var sheet = workbook.Worksheet(1);
                        for (int i = 2; i < 500; i++)
                        {
                            if (sheet.Cell(i, 1).IsEmpty() != true)
                            {
                                SymbolList.Add(sheet.Cell(i, 1).Value.ToString());
                            }
                            else { break; }
                        }
                    }
                    break;
                }
                catch
                {
                    Console.WriteLine(" Не смог открыть ексель Setting.xlsx в папке WorkMexc\n" +
                                      " Проверь не открыта ли ексель или есть ли доступ");
                    Thread.Sleep(10000);
                }
            }
        }
        public static async Task StartNewExelAsync(BybitRestClient bybitRestClient)
        {
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.WriteLine();
            Console.WriteLine(" Хотите создать новую ексель для торговой пары?\n" +
                              " Напишите ДА или НЕТ и нижмите ENTER");
            while (true)
            {
                string response = Console.ReadLine();

                //Создание ексель под торговую пару
                if (response.ToUpper() == "ДА")
                {
                    Console.WriteLine(" Укажите торговую пару и нажмите ENTER\n" +
                                      " Пример: BTCUSDT");
                    string tradingPair = Console.ReadLine();

                    //Код для создания новой ексель под торговую пару
                    await NewExcel.TradingPairAsync(tradingPair.ToUpper(), bybitRestClient);

                    Console.WriteLine(" Хотите ещё создать ексель для торговой пары?\n" +
                                      " Напишите ДА или НЕТ и нижмите ENTER");
                }
                else { break; }
            }
            Console.WriteLine(" Создать ли Вам сетку для торговой пары?\n" +
                              " Напишите ДА или НЕТ и нижмите ENTER");
            while (true)
            {
                string response = Console.ReadLine();

                //Создание ексель под торговую пару
                if (response.ToUpper() == "ДА")
                {
                    Console.WriteLine(" Укажите торговую пару и нажмите ENTER\n" +
                                      " Пример: BTCUSDT");
                    string tradingPair = Console.ReadLine();

                    //Код для создания сетки под торговую пару
                    await NewExcel.Setka(tradingPair.ToUpper(), bybitRestClient);

                    Console.WriteLine(" Хотите ещё создать сетку для торговой пары?\n" +
                                      " Напишите ДА или НЕТ и нижмите ENTER");
                }
                else { break; }
            }
        }
        public static async Task StartNewExelAsyncMexc(MexcRestClient mexcRestClient)
        {
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.WriteLine();
            Console.WriteLine(" Хотите создать новую ексель для торговой пары?\n" +
                              " Напишите ДА или НЕТ и нижмите ENTER");
            while (true)
            {
                string response = Console.ReadLine();

                //Создание ексель под торговую пару
                if (response.ToUpper() == "ДА")
                {
                    Console.WriteLine(" Укажите торговую пару и нажмите ENTER\n" +
                                      " Пример: BTCUSDT");
                    string tradingPair = Console.ReadLine();

                    //Код для создания новой ексель под торговую пару
                    await NewExcel.TradingPairAsyncMexc(tradingPair.ToUpper(), mexcRestClient);

                    Console.WriteLine(" Хотите ещё создать ексель для торговой пары?\n" +
                                      " Напишите ДА или НЕТ и нижмите ENTER");
                }
                else { break; }
            }
            Console.WriteLine(" Создать ли Вам сетку для торговой пары?\n" +
                              " Напишите ДА или НЕТ и нижмите ENTER");
            while (true)
            {
                string response = Console.ReadLine();

                //Создание ексель под торговую пару
                if (response.ToUpper() == "ДА")
                {
                    Console.WriteLine(" Укажите торговую пару и нажмите ENTER\n" +
                                      " Пример: BTCUSDT");
                    string tradingPair = Console.ReadLine();

                    //Код для создания сетки под торговую пару
                    await NewExcel.SetkaMexc(tradingPair.ToUpper(), mexcRestClient);

                    Console.WriteLine(" Хотите ещё создать сетку для торговой пары?\n" +
                                      " Напишите ДА или НЕТ и нижмите ENTER");
                }
                else { break; }
            }
        }
    }
}
