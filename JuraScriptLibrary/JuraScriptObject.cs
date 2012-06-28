using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Jurassic;
using JuraScriptLibrary.Require;
using JuraScriptLibrary.COM;
using System.IO;
using Jurassic.Library;

namespace JuraScriptLibrary
{
    public class JuraScriptObject : IJuraScriptHost
    {
        /// <summary>
        /// Get or set the global JavaScript variable 'exports'
        /// </summary>
        public object Exports
        {
            get
            {
                return ScriptEngine.GetGlobalValue("exports");
            }
            set
            {
                ScriptEngine.SetGlobalValue("exports", value);
            }
        }
        string[] _Args = null;
        ScriptEngine ScriptEngine = new ScriptEngine();
        public JuraScriptObject()
        {
            Initialize();
        }

        private void Initialize()
        {
            ScriptEngine.EnableExposedClrTypes = true;
            WScriptWrapper wsw = new WScriptWrapper(_Args, "CODEME_UNKNOWN", this);
            ScriptEngine.SetGlobalValue("WScript", wsw);

            ScriptEngine.SetGlobalValue("ActiveXObject", new ActiveXConstructor(ScriptEngine));
            ScriptEngine.SetGlobalValue("Enumerator", new EnumeratorConstructor(ScriptEngine));
            ScriptEngine.SetGlobalValue("Console", typeof(JSConsole));
            ScriptEngine.SetGlobalValue("require", new requireObject(JurassicEngine.Object.InstancePrototype));
            ScriptEngine.SetGlobalValue("exports", new exportsObject(JurassicEngine.Object.InstancePrototype));
            

        }

        public void StartUp(string[] args)
        {
            _Args = args;
            if (args.Length > 0)
            {
                if (args.FirstOrDefault(a => a.ToUpper().StartsWith("/CON")) != null)
                {
                    ExecuteShell(args);
                }
                if (File.Exists(args[0]))
                {
                    ExecuteCommandline(args);
                }
            }
        }

        public void Execute(string code)
        {
            ScriptEngine.Execute(code);
        }

        public void ExecuteCommandline(string[] args)
        {
            ScriptEngine.ExecuteFile(args[0]);
        }

        public void ExecuteShell(string[] args)
        {
            Console.WriteLine("JuraScript -- Jurassic JavaScript Shell");
            Console.WriteLine("");

            do
            {
                Console.Write("jur>");
                string cmdStr = Console.ReadLine();
                if (new string[] { "EXIT", "QUIT", "BYE" }.Contains(cmdStr.ToUpper()))
                { break; }
                else
                {
                    try
                    {
                        ScriptEngine.Execute(cmdStr);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error: " + ex.ToString());
                    }
                }
            } while (true);
        }

        public ScriptEngine JurassicEngine
        {
            get
            {
                return ScriptEngine;
            }
            set
            {
                ScriptEngine = value;
            }
        }

        public void Quit()
        {
            Environment.Exit(0);
        }

        public void Quit(int errCode)
        {
            Environment.Exit(errCode);
        }
    }
}
