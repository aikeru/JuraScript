using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace JuraScriptLibrary
{
    public class WScriptWrapper
    {
        public WScriptWrapper()
        {
        }
        public void Echo(object txt)
        {
            Console.WriteLine(txt.ToString());
        }
        public string Read()
        {
            return ((Char)Console.Read()).ToString();
        }
    }
}
