using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Jurassic.Library;
using Jurassic;
using System.IO;

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
                && argumentValues.Length > 0
                && argumentValues[0] != null
                && argumentValues[0].GetType() == typeof(String)
                && !String.IsNullOrWhiteSpace(argumentValues[0].ToString()))
            {
                string mod = argumentValues[0].ToString();
                if (File.Exists(mod + ".jur"))
                {
                    mod = mod + ".jur";
                }
                
                if (!mod.ToUpper().EndsWith(".JUR"))
                {
                    if (File.Exists(mod + ".jur")) { mod = mod + ".jur"; }
                }

                JuraScriptObject jso = new JuraScriptObject();
                jso.ExecuteCommandline(new string[] { mod });

                return jso.Exports;    
            } else {
                throw new JavaScriptException(this.Engine, "Invalid argument", "Must pass a string to 'require()'.");
            }
            return null;
        }
    }
}
