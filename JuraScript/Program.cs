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
            JuraScriptObject jso = new JuraScriptObject();
            jso.StartUp(args);
        }
    }
}
