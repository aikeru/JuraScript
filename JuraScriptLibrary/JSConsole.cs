using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using JurassicTest;

namespace JuraScriptLibrary
{

    public class JSConsole
    {
        public static void WriteLine(object arg)
        {
            Console.WriteLine(arg);
        }
        public static string ReadLine()
        {
            return Console.ReadLine();
        }
    }
}
