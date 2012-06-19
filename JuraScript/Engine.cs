using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Jurassic;
using System.IO;
using JuraScriptLibrary;
using System.Threading;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Reflection;
using JuraScript;
using JuraScriptLibrary.COM;
using JuraScriptLibrary;

namespace JuraScriptLibrary {
    //// {F44625AB-0BAA-4F5E-9693-8CA1CE0C1558}
    //IMPLEMENT_OLECREATE(<<class>>, <<external_name>>, 
    //0xf44625ab, 0xbaa, 0x4f5e, 0x96, 0x93, 0x8c, 0xa1, 0xce, 0xc, 0x15, 0x58);
    //LScript was 62416980-8B84-4c51-A09C-AB9623C3CC3E
    //JuraScript is F44625AB-0BAA-4F5E-9693-8CA1CE0C1558
    //LScript ProgId was "LScript"
    //JuraScript ProgId is "JuraScript"

    [ClassInterface(ClassInterfaceType.None),
     ComVisible(true),
     //Guid("k"),
     Guid("F44625AB-0BAA-4F5E-9693-8CA1CE0C1558"), //msnead - 2012 picking this back up put this here for now
     ProgId("JuraScript")]
    public class Engine : IActiveScript, IActiveScriptParse, IDisposable //, IObjectSafety
        , IJuraScriptHost {
        TraceListener traceListener = new TextWriterTraceListener(Console.Out);
        IActiveScriptSite site;
        ScriptState currentScriptState = ScriptState.Uninitialized;
        ScriptEngine Jurassic = null;

        public ScriptEngine JurassicEngine {
            get { return Jurassic; }
            set { Jurassic = value; }
        }

        public Engine() {
            Debug.Listeners.Add(traceListener);
        }

        ~Engine() {
            Dispose(false);
        }

        public void Dispose() {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing) {
            if (disposing) {
                if (traceListener != null) {
                    Debug.Listeners.Remove(traceListener);
                    traceListener.Dispose();
                    traceListener = null;
                }
            }
        }

        public void SetScriptSite(IActiveScriptSite site) {
            Debug.WriteLine("SetScriptSite()");
            this.site = site;
        }

        public void GetScriptSite(ref Guid riid, out IntPtr site) {
            Debug.WriteLine("GetScriptSite()");
            site = IntPtr.Zero;
        }

        public void SetScriptState(ScriptState state) {
            Debug.WriteLine("SetScriptState(" + state.ToString() + ")");
            this.currentScriptState = state;
            // site.OnStateChange (this.currentScriptState);  
        }

        public void GetScriptState(out ScriptState scriptState) {
            Debug.WriteLine("GetScriptState()");
            scriptState = this.currentScriptState;
        }

        public void Close() {
            Debug.WriteLine("Close()");
        }

        void ProcessConstructor(ConstructorInfo constructorInfo) {
            Debug.WriteLine("Add constructor " + constructorInfo.ToString());
        }

        void ProcessEvent(EventInfo eventInfo) {
            Debug.WriteLine("Add event " + eventInfo.ToString());
        }

        void ProcessMethod(System.Reflection.MethodInfo methodInfo) {
            Debug.WriteLine("Add method " + methodInfo.ToString());
        }

        void ProcessProperty(PropertyInfo propertyInfo) {
            Debug.WriteLine("Add property " + propertyInfo.ToString());
        }

        void ProcessMember(System.Reflection.MemberInfo memInfo) {
            System.Reflection.MethodInfo methodInfo = memInfo as System.Reflection.MethodInfo;
            if (methodInfo != null)
                ProcessMethod(methodInfo);
            else {
                System.Reflection.ConstructorInfo constructorInfo = memInfo as ConstructorInfo;
                if (constructorInfo != null)
                    ProcessConstructor(constructorInfo);
                else {
                    PropertyInfo propertyInfo = memInfo as PropertyInfo;
                    if (propertyInfo != null)
                        ProcessProperty(propertyInfo);
                    else {
                        EventInfo eventInfo = memInfo as EventInfo;
                        if (eventInfo != null)
                            ProcessEvent(eventInfo);
                        else
                            Debug.WriteLine("Add member " + memInfo.ToString());
                    }
                }
            }
        }

        void ProcessItem(string name, ScriptItem flags, object item, Type itemType) {
            if ((flags & ScriptItem.IsVisible) != 0) {
                //CL.SetSymbolValue(CL.Intern(name), item);
            }
            if ((flags & ScriptItem.GlobalMembers) != 0) {
                foreach (FieldInfo fieldInfo in itemType.GetFields()) {
                    Debug.WriteLine("Process Field " + fieldInfo.ToString());
                }
                foreach (PropertyInfo propertyInfo in itemType.GetProperties()) {
                    string propName = propertyInfo.Name;
                    //CL.SetSymbolValue(CL.Intern(propName), propertyInfo);
                }
                Debug.WriteLine("Installed");
            }
            Debug.WriteLine("processed");
        }

        public void AddNamedItem(string name, ScriptItem flags) {
            Debug.WriteLine("AddNamedItem(\"" + name + "\", " + flags.ToString() + ")");
            // What did he give us?
            object item;
            object typeinfo;
            if (this.site != null) {
                this.site.GetItemInfo(name, ScriptInfo.IUnknown, out item, out typeinfo);
                Type itemType = item.GetType();
                if (itemType.FullName == "System.__ComObject") {
                    throw new NotImplementedException();
                } else {
                    ProcessItem(name, flags, item, itemType);
                }

            } else {
                throw new NotImplementedException();
            }

            //if (this.site != null) {
            //    try {
            //        this.site.GetItemInfo (name, ScriptInfo.IUnknown, out item, out typeinfo);
            //        Debug.WriteLine ("item is " + item.ToString ());
            //        Type type = (item as IDispatch).GetTypeInfo (0, 0x409);
            //        Debug.WriteLine ("Type is " + type.ToString ());
            //        foreach (System.Reflection.MemberInfo minfo in type.GetMembers ())
            //            Debug.WriteLine (minfo);
            //    }
            //    catch (Exception e) {
            //        Debug.WriteLine ("Exception " + e.ToString ());
            //    }
            //}
            //else {
            //    // I guess we can find out later.
            //    Debug.WriteLine ("No site, but named items being added.");
            //}
        }

        public void AddTypeLib(ref Guid typeLib, uint major, uint minor, uint flags) {
            Debug.WriteLine("AddTypeLib()");
        }

        public void GetScriptDispatch(string itemName, out object dispatch) {
            Debug.WriteLine("GetScriptDispatch()");
            dispatch = null;
        }
        public void GetCurrentScriptThreadID(out uint thread) {
            Debug.WriteLine("GetCurrentScriptThreadID()");
            thread = 0;
        }
        public void GetScriptThreadID(uint win32ThreadId, out uint thread) {
            Debug.WriteLine("GetScriptThreadID()");
            thread = 0;
        }
        public void GetScriptThreadState(uint thread, out ScriptThreadState state) {
            Debug.WriteLine("GetScriptThreadState()");
            state = ScriptThreadState.NotInScript;
        }
        public void InterruptScriptThread(uint thread, IntPtr excepinfo, uint flags) {
            Debug.WriteLine("InterruptScriptThread()");
        }

        public void Clone(out IActiveScript script) {
            Debug.WriteLine("Clone()");
            script = this;
        }

        public void InitNew() {
            Debug.WriteLine("InitNew()");
            //DotNet.Enable();
            //CLOS.Init();
            this.currentScriptState = ScriptState.Disconnected;
            if (this.site != null)
                this.site.OnStateChange(this.currentScriptState);
            //CL.Readtable.Case = ReadtableCase.Preserve;

            //ms- I guess this is where LScript would start up a new instance of the engine
            WScriptWrapper wsw = new WScriptWrapper(new string[] { }, "UNKNOWN_CODEME", this);
            Jurassic = new ScriptEngine();
            Jurassic.SetGlobalValue("ActiveXObject", new ActiveXConstructor(Jurassic));
            Jurassic.SetGlobalValue("Enumerator", new EnumeratorConstructor(Jurassic));
            Jurassic.SetGlobalValue("Console", typeof(JSConsole));
            Jurassic.SetGlobalValue("WScript", wsw);
        }

        public void AddScriptlet(string defaultName,
                    string code,
                    string itemName,
                    string subItemName,
                    string eventName,
                    string delimiter,
                    uint sourceContextCookie,
                    uint startingLineNumber,
                    uint flags,
                    IntPtr name,
                    IntPtr info) {
            Debug.WriteLine("AddScriptlet");
        }

        public void ParseScriptText(
                string code,
                string itemName,
                object context,
                string delimiter,
                int sourceContextCookie,
                uint startingLineNumber,
                ScriptText flags,
                IntPtr result,
                IntPtr info) {
            Debug.WriteLine("ParseScriptText()");
            Debug.WriteLine("code: " + code);
            Debug.WriteLine("itemName: " + itemName);
            Debug.WriteLine("delimiter: " + delimiter);
            Debug.WriteLine("flags: " + flags);
            Debug.Flush();

            ////result = null;
            //// info = new System.Runtime.InteropServices.ComTypes.EXCEPINFO ();
            //CL.Eval(CL.ReadFromString(code));

            //ms- I think this is where LScript processes a block of script, so let's do that
            Jurassic.Execute(code);
        }

        //private const int INTERFACESAFE_FOR_UNTRUSTED_CALLER = 0x00000001;
        //private const int INTERFACESAFE_FOR_UNTRUSTED_DATA = 0x00000002;
        //private const int S_OK = 0;

        //public int GetInterfaceSafetyOptions (ref Guid riid,
        //         out int pdwSupportedOptions,
        //         out int pdwEnabledOptions)
        //{
        //    pdwSupportedOptions = 3;
        //    pdwEnabledOptions = 3;
        //    return S_OK;
        //}

        //public int SetInterfaceSafetyOptions (ref Guid riid, 
        //    int dwOptionSetMask, 
        //    int dwEnabledOptions)
        //{
        //    return S_OK;
        //}

        #region IJuraScriptHost Members

        public void Quit() {
            //Don't know how to do this yet...
        }

        public void Quit(int errCode) {
            //See Quit()
        }

        #endregion
    }


    public class JuraScriptObject : IJuraScriptHost {
        ScriptEngine ScriptEngine = new ScriptEngine();

        public ScriptEngine JurassicEngine {
            get { return ScriptEngine; }
            set { ScriptEngine = value; }
        }

        public JuraScriptObject() {
        }
        public void ExecuteCommandline(string[] args) {
            WScriptWrapper wsw = new WScriptWrapper(args, args[0], this);

            ScriptEngine.SetGlobalValue("ActiveXObject", new ActiveXConstructor(ScriptEngine));

            ScriptEngine.SetGlobalValue("Enumerator", new EnumeratorConstructor(ScriptEngine));

            ScriptEngine.SetGlobalValue("Console", typeof(JSConsole));

            ScriptEngine.SetGlobalValue("WScript", wsw);

            if (args.Length > 0) {
                if (File.Exists(args[0])) {

                    ScriptEngine.ExecuteFile(args[0]);

                }
            }
        }

        #region IJuraScriptHost Members

        public void Quit() {
            Environment.Exit(0);
        }

        public void Quit(int errCode) {
            Environment.Exit(errCode);
        }

        #endregion
    }
}
