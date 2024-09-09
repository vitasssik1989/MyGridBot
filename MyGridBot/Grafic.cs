using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyGridBot
{
    internal class Grafic
    {
        static public void Write(int i)
        {
            if (i < 1000)
            {
                Console.WriteLine(" ■");
            }
            else if (i > 1000 && i < 2000)
            {
                Console.WriteLine(" ■ ■");
            }
            else if (i > 2000 && i < 3000)
            {
                Console.WriteLine(" ■ ■ ■");
            }
            else if (i > 3000 && i < 4000)
            {
                Console.WriteLine(" ■ ■ ■ ■");
            }
            else
            {
                Console.WriteLine(" ■ ■ ■ ■ ■");
            }
        }

        static public void GreetUser()
        {
            Console.ForegroundColor = ConsoleColor.Blue;
            Random rnd = new Random();
            string hello = " ...Проект BOVI предоставляет уникальную возможность\n" +
                           " пользоваться бесплатным,автоматизированным\n" +
                           " многофункциональным GRID-ботом,\n" +
                           " тщательно настроенным для\n" +
                           " потребностей трейдеров на площадке BYBIT и MEXC. \n\n" +
                           " ...За каждый донат, участник получает BOVI токены,\n" +
                           " средства от которых направляются\n" +
                           " на улучшение и оптимизацию работы основного бота\n" +
                           " на одном торговом аккаунте.\n\n" +
                           " ...Держатели токенов будут впоследствии участвовать\n" +
                           " в общей прибыли от основного бота.\n\n" +
                           " ...Приветсвуется любая помощь в продвижении проекта\n\n" +
                           " ...Желаю тебе больше профита в мире крипты ;)\n\n" +
                           "                           Разработчик: @Vitasssik\n" +
                           " Нажмите ENTER";

            Console.Write(hello);
            Console.ReadLine();
        }
    }
}
