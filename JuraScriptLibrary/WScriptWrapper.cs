using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using JurassicTest;
using System.Collections;
using System.Threading;
using Jurassic.Library;
using Jurassic;

namespace JuraScriptLibrary {
    public interface IJuraScriptHost {
        ScriptEngine JurassicEngine { get; set;}
        void Quit();
        void Quit(int errCode);
    }
    //http://msdn.microsoft.com/en-us/library/312a5kbt(v=vs.85).aspx TextStream
    //http://msdn.microsoft.com/en-us/library/1y8934a7(v=vs.85).aspx StdIn
    public class WScriptStdIn {

        private ScriptEngine _Engine = null;

        public WScriptStdIn(ScriptEngine engine) {
            _Engine = engine;
        }

        public string Read() {
            return ((char)Console.Read()).ToString();
        }
        public string Read(int characters) {
            string retStr = "";
            for (var i = 0; i < characters; i++) {
                retStr += ((char)Console.Read()).ToString();
            }
            return retStr;
        }

        public string ReadAll() {
            throw new JavaScriptException(_Engine, "JuraScriptNotImplemented", "Not yet implemented.");
        }

        public string ReadLine() {
            return Console.ReadLine();
        }

        public void Skip(int characters) {
            for (var i = 0; i < characters; i++) {
                Console.Read();
            }
        }
        public void SkipLine() {
            Console.ReadLine();
        }
        //CScript Read-Only Property
        public bool AtEndOfLine {
            get { throw new JavaScriptException(_Engine, "JuraScriptNotImplemented", "Not yet implemented."); }
        }
        //CScript Read-Only Property
        public int Column {
            get { throw new JavaScriptException(_Engine, "JuraScriptNotImplemented", "Not yet implemented."); }
        }
        //CScript Read-Only Property
        public int Line {
            get { throw new JavaScriptException(_Engine, "JuraScriptNotImplemented", "Not yet implemented."); }
        }
        //http://msdn.microsoft.com/en-us/library/dxkz5tcx(v=vs.85).aspx
        public void Close() {
            throw new JavaScriptException(_Engine, "WScriptError", "Not applicable.");
        }
    }
    //http://msdn.microsoft.com/en-us/library/312a5kbt(v=vs.85).aspx
    public class WScriptStdErr {
        public WScriptStdErr(ScriptEngine engine) {
        }
    }
    //http://msdn.microsoft.com/en-us/library/312a5kbt(v=vs.85).aspx
    public class WScriptStdOut {
        public WScriptStdOut(ScriptEngine engine) {
        }
    }

    public class WScriptWrapper {
        private IJuraScriptHost _Host = null;

        string[] _Arguments = new string[] { };
        public WScriptWrapper(string[] args, string scriptFullName, IJuraScriptHost host) {
            _ScriptFullName = scriptFullName;
            _Arguments = args;
            _Host = host;
            StdErr = new WScriptStdErr(_Host.JurassicEngine);
            StdIn = new WScriptStdIn(_Host.JurassicEngine);
            StdOut = new WScriptStdOut(_Host.JurassicEngine);
        }

        private string GetObjectEcho(object txt) {
            string objStr = "";
            IEnumerable txtEnum = null;
            if (txt.GetType() == typeof(ArrayInstance)) {
                txtEnum = (txt as Jurassic.Library.ArrayInstance).ElementValues;
            } else { txtEnum = txt as IEnumerable; }

            if (txtEnum != null) {
                foreach (var x in txtEnum) {
                    objStr += String.IsNullOrEmpty(objStr) ? GetObjectEcho(x) : "," + GetObjectEcho(x);
                    //objStr += String.IsNullOrEmpty(objStr) ? x.ToString() : "," + x.ToString();
                }
            } else {
                objStr = txt.ToString();
            }
            return objStr;
        }
        /// <summary>
        /// Attempts to mimic the functionality observed in WSH
        /// </summary>
        /// <param name="txt"></param>
        public void Echo(object txt) {
            Console.WriteLine(GetObjectEcho(txt));
        }

        public string Read() {
            return ((Char)Console.Read()).ToString();
        }

        public string[] Arguments {
            get {
                return _Arguments;
            }
        }

        public string Version {
            get {
                return "0";
            }
        }
        public string BuildVersion {
            get {
                return "";
            }
        }
        public string FullName {
            get {
                return Assembly.GetEntryAssembly().FullName;
            }
        }
        private bool _Interactive = true;
        public bool Interactive {
            get {
                return _Interactive;
            }
            set {
                _Interactive = value;
            }
        }
        public string Name {
            get {
                return "Jurassic Script Host";
            }
        }
        public string Path {
            get {
                return Assembly.GetExecutingAssembly().Location;
            }
        }
        private string _ScriptFullName = "";

        /// <summary>
        /// Returns the full path of the currently running script
        /// </summary>
        public string ScriptFullName {
            get {
                return _ScriptFullName;
            }
        }
        private string _ScriptName = "";
        /// <summary>
        /// Returns the name of the script being run
        /// </summary>
        public string ScriptName {
            get {
                return _ScriptName;
            }
        }

        //StdErr property, a TextStream Object
        //http://msdn.microsoft.com/en-us/library/hyez2k48(v=vs.85).aspx

        public WScriptStdIn StdIn = null;

        public WScriptStdOut StdOut = null;

        public WScriptStdErr StdErr = null;

        /// <summary>
        /// Connects the object's event sources to functions with a given prefix.<br/>
        /// http://msdn.microsoft.com/en-us/library/ccxe1xe6(v=vs.85).aspx
        /// </summary>
        /// <param name="objEventSource"></param>
        /// <param name="strPrefix"></param>
        /// <returns></returns>
        public object ConnectObject(object objEventSource, string strPrefix) {
            throw new NotImplementedException();
        }
        //http://msdn.microsoft.com/en-us/library/xzysf6hc(v=vs.85).aspx
        public object CreateObject(string progID, string strPrefix) {
            throw new NotImplementedException();
        }
        public object CreateObject(string progID) {
            return new ActiveXInstance(null, progID);
        }
        //http://msdn.microsoft.com/en-us/library/2d26y0c1(v=vs.85).aspx
        public object DisconnectObject(object obj) {
            throw new NotImplementedException();
        }
        //http://msdn.microsoft.com/en-us/library/8ywk619w(v=vs.85).aspx
        public object GetObject(string pathName) {
            return new NotImplementedException();
        }
        public object GetObject(string pathName, string progID) {
            return new NotImplementedException();
        }
        public object GetObject(string pathName, string progID, string prefix) {
            return new NotImplementedException();
        }
        public void Quit() {
            _Host.Quit();
        }
        public void Quit(int errCode) {
            _Host.Quit(errCode);
        }
        public void Sleep(int time) {
            Thread.Sleep(time);
        }
    }
}
