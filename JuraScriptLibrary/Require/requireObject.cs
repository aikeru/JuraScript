using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Jurassic.Library;
using Jurassic;

namespace JuraScriptLibrary.Require
{
    public class requireObject : FunctionInstance
    {
        public requireObject(ObjectInstance prototype) : base(prototype)
        {
            this.PopulateFunctions();
        }
        public override object CallLateBound(object thisObject, params object[] argumentValues)
        {
            if (argumentValues != null
                && argumentValues.Length > 0)
            {
                var mod = argumentValues[0];
                if (mod.GetType() == typeof(String))
                {
                    //Find module

                }
                Console.WriteLine("TYPE OF ARGUMENT: " + mod.GetType());
                Console.WriteLine(mod as ArrayInstance != null ? "It's an array" : "It's not an array");
                Console.WriteLine(mod as StringInstance != null ? "It's a string" : "It's not a string");
            }
            return null;
        }
    }
}
