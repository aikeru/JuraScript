using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Jurassic.Library;
using Jurassic;
using System.IO;
using System.Reflection;

namespace JuraScriptLibrary.Require
{
    public class requireObject : FunctionInstance
    {
        public static Dictionary<string, string> BuiltInModules = new Dictionary<string, string>()
        {
            { "assert", "JuraScriptLibrary.Scripts.commonjs_assert.js" }
        };
        private JuraScriptObject LoadBuiltInResource(string resName)
        {
            JuraScriptObject jso = new JuraScriptObject();

            Stream s = Assembly.GetAssembly(typeof(requireObject)).GetManifestResourceStream(BuiltInModules[resName]);
            StreamReader sr = new StreamReader(s);
            jso.Execute(sr.ReadToEnd());
            sr.Close();
            s.Close();

            return jso;
        }

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
                JuraScriptObject jso = null;

                if (BuiltInModules.ContainsKey(mod))
                {
                    jso = LoadBuiltInResource(mod);
                }
                else
                {

                    if (File.Exists(mod + ".jur"))
                    {
                        mod = mod + ".jur";
                    }

                    if (!mod.ToUpper().EndsWith(".JUR"))
                    {
                        if (File.Exists(mod + ".jur")) { mod = mod + ".jur"; }
                    }

                    jso = new JuraScriptObject();
                    jso.ExecuteCommandline(new string[] { mod });
                }

                return jso.Exports;    
            } else {
                throw new JavaScriptException(this.Engine, "Invalid argument", "Must pass a string to 'require()'.");
            }
            return null;
        }
    }
}
