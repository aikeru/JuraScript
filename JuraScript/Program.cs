using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Jurassic;
using System.IO;
using JuraScriptLibrary;
using JuraScriptLibrary.COM;
using JuraScriptLibrary.Require;

namespace JuraScript
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0
                || args[0] == "/?" || args[0] == "--?")
            {
                Console.WriteLine();
                Console.WriteLine("JuraScript");
                Console.WriteLine("(c) " + DateTime.Now.Year.ToString() + " Michael Snead");
                Console.WriteLine();
                Console.WriteLine("Usage: ");
                Console.WriteLine();
                Console.WriteLine(" jurascript /CON - Start JuraScript console/shell");
                Console.WriteLine(" jurascript <filename> - Execute script <filename>");
                Console.WriteLine();
            }
            else
            {
                JuraScriptObject jso = new JuraScriptObject();
                jso.StartUp(args);
            }
        }
    }
}
