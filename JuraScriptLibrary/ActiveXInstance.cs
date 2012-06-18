using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Jurassic.Library;
using Jurassic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using JuraScriptLibrary;

namespace JurassicTest
{

    public class ActiveXFunction : FunctionInstance
    {
        private string ProgID = "";
        private ActiveXInstance _ActiveXObjectInstance = null;
        public ActiveXInstance ActiveXObjectInstance
        {
            get
            {
                return _ActiveXObjectInstance;
            }
            set
            {
                _ActiveXObjectInstance = value;
            }
        }
        private Type _ActiveXType = null;
        private string _PropertyName = "";
        public string PropertyName
        {
            get
            {
                return _PropertyName;
            }
            set
            {
                _PropertyName = value;
            }
        }
        public ActiveXFunction(ObjectInstance prototype, ActiveXInstance instanceobject, string propertyName)
            : base(prototype)
        {
            ActiveXObjectInstance = instanceobject;
            _PropertyName = propertyName;
            this.PopulateFunctions();
        }

        public object ActiveXUnpack()
        {
            try
            {
                return ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(_PropertyName,
                    BindingFlags.GetProperty, null, ActiveXObjectInstance.ActiveXObject, null);
            }
            catch (Exception ex)
            {
                return ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(_PropertyName,
                    BindingFlags.InvokeMethod, null, ActiveXObjectInstance.ActiveXObject, null);
            }
        }

        private object ActiveXGetValue(string propertyName)
        {

            JuraScriptLibrary.ComReflector.ComReflectedMemberTypes memberType = ComReflector.GetMemberType(ActiveXObjectInstance.ActiveXObject, _PropertyName);
            if (memberType == ComReflector.ComReflectedMemberTypes.Method)
            {
                //Get the property off of the .. method ?? Most likely never happen.
                var propertyValue = new ActiveXInstance(this.Prototype, ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(_PropertyName,
                        BindingFlags.InvokeMethod, null, ActiveXObjectInstance.ActiveXObject, null));
                return new ActiveXFunction(this.Prototype, propertyValue, propertyName);
            }
            else if (memberType != ComReflector.ComReflectedMemberTypes.NOT_FOUND && memberType != ComReflector.ComReflectedMemberTypes.NO_ITYPEINFO)
            {
                //Unwrap the ActiveXInstance ActiveX Object, then create a new ActiveXFunction pointing at the new property
                var propertyValue = new ActiveXInstance(this.Prototype, ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(_PropertyName,
                    BindingFlags.GetProperty, null, ActiveXObjectInstance.ActiveXObject, null));
                return new ActiveXFunction(this.Prototype, propertyValue, propertyName);
            }

            throw new JavaScriptException(this.Engine, "COMError", "No such property.");

            /* OLD METHOD
            try
            {
                //Unwrap the ActiveXInstance ActiveX Object, then create a new ActiveXFunction pointing at the new property
                var propertyValue = new ActiveXInstance(this.Prototype, ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(_PropertyName,
                    BindingFlags.GetProperty, null, ActiveXObjectInstance.ActiveXObject, null));
                return new ActiveXFunction(this.Prototype, propertyValue, propertyName);
            }
            catch (Exception ex)
            {
                try
                {
                    //Get the property off of the .. method ?? Most likely never happen.
                    var propertyValue = new ActiveXInstance(this.Prototype, ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(_PropertyName,
                        BindingFlags.InvokeMethod, null, ActiveXObjectInstance.ActiveXObject, null));
                    return new ActiveXFunction(this.Prototype, propertyValue, propertyName);
                }
                catch (Exception ex2)
                {
                    throw new JavaScriptException(Engine, "COMError", ex2.ToString());
                }
            }
            
            return base.GetMissingPropertyValue(propertyName); */
        }

        protected override object GetMissingPropertyValue(string propertyName)
        {
            return ActiveXGetValue(propertyName);
        }

        public override object CallLateBound(object thisObject, params object[] argumentValues)
        {
            //if ((thisObject is ActiveXFunction) == false)
            //    throw new JavaScriptException(this.Engine, "TypeError", "Invalid 'this' value.");
            //return ((ActiveXFunction)thisObject).CallFunction(this._PropertyName, argumentValues);
            return CallFunction(this._PropertyName, argumentValues);
        }

        private object CallFunction(string functionName, params object[] argumentValues)
        {
            //return functionName + argumentValues.Length;
            //return "somepropertyok";

            /************** OLD METHOD ****************
            try
            {
                //unpack it first
                try
                {
                    var propertyValue = ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(_PropertyName,
                                        BindingFlags.InvokeMethod, null, ActiveXObjectInstance.ActiveXObject, argumentValues);
                    if (propertyValue == null) { return null; }
                    return new ActiveXInstance(this.Prototype, propertyValue);
                    //return new ActiveXInstance(this.Prototype, propertyValue.GetType().InvokeMember(functionName,
                    //                BindingFlags.InvokeMethod, null, propertyValue, argumentValues));
                }
                catch (Exception unpackEx)
                {
                    //Sometimes CallFunction is getting a property
                    //ie: Excel .Cells(x,x) which is an array but is called like a function even in JScript
                    // (see MSDN examples and test it in WSH)
                    var propertyValue = ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(_PropertyName,
                                        BindingFlags.GetProperty, null, ActiveXObjectInstance.ActiveXObject, null);
                    if (propertyValue == null) { return null; }
                    return new ActiveXInstance(this.Prototype,
                        propertyValue.GetType().InvokeMember(functionName, BindingFlags.GetProperty, null, propertyValue, argumentValues));
                }

                
            }
            catch (Exception ex)
            {
                throw new JavaScriptException(this.Engine, "COMError", ex.ToString());
            }
            throw new JavaScriptException(this.Engine, "COMError", "No such method.");
            **********************************************/

            //Is it a property or a method?
            JuraScriptLibrary.ComReflector.ComReflectedMemberTypes memberType = ComReflector.GetMemberType(ActiveXObjectInstance.ActiveXObject, _PropertyName);
            if(memberType == ComReflector.ComReflectedMemberTypes.Method) {
                var memberValue = ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(_PropertyName,
                                        BindingFlags.InvokeMethod, null, ActiveXObjectInstance.ActiveXObject, argumentValues);
                if(memberValue == null) { return null; }
                return new ActiveXInstance(this.Prototype, memberValue);
            } else if(memberType != ComReflector.ComReflectedMemberTypes.NOT_FOUND && memberType != ComReflector.ComReflectedMemberTypes.NO_ITYPEINFO) {
                //Sometimes CallFunction is getting a property
                //ie: Excel .Cells(x,x) which is an array but is called like a function even in JScript
                var memberValue = ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(_PropertyName,
                            BindingFlags.GetProperty, null, ActiveXObjectInstance.ActiveXObject, null);
                if (memberValue == null) { return null; }
                return new ActiveXInstance(this.Prototype,
                        memberValue.GetType().InvokeMember(functionName, BindingFlags.GetProperty, null, memberValue, argumentValues));
            }

            throw new JavaScriptException(this.Engine, "COMError", "No such method.");
        }

        /// <summary>
        /// Returns a primitive value associated with the object.
        /// </summary>
        /// <returns> A primitive value associated with the object. </returns>
        [JSFunction(Name = "valueOf")]
        public new ObjectInstance ValueOf()
        {

            JuraScriptLibrary.ComReflector.ComReflectedMemberTypes memberType = ComReflector.GetMemberType(ActiveXObjectInstance.ActiveXObject, _PropertyName);
            if (memberType == ComReflector.ComReflectedMemberTypes.Method)
            {
                //Get the property off of the .. method ?? Most likely never happen.
                var propertyValue = new ActiveXInstance(this.Prototype, ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(_PropertyName,
                        BindingFlags.InvokeMethod, null, ActiveXObjectInstance.ActiveXObject, null));
                return new ActiveXInstance(this.Prototype, propertyValue);
            }
            else if (memberType != ComReflector.ComReflectedMemberTypes.NOT_FOUND && memberType != ComReflector.ComReflectedMemberTypes.NO_ITYPEINFO)
            {
                //Unwrap the ActiveXInstance ActiveX Object, then create a new ActiveXFunction pointing at the new property
                var propertyValue = new ActiveXInstance(this.Prototype, ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(_PropertyName,
                    BindingFlags.GetProperty, null, ActiveXObjectInstance.ActiveXObject, null));
                return new ActiveXInstance(this.Prototype, propertyValue);
            }

            throw new JavaScriptException(this.Engine, "COMError", "Error evaluating: " + _PropertyName);


            /* OLD METHOD
            try
            {
                var propertyValue = ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(_PropertyName,
                    BindingFlags.GetProperty, null, ActiveXObjectInstance.ActiveXObject, null);
                return new ActiveXInstance(this.Prototype, propertyValue);
            }
            catch (Exception ex)
            {
                try
                {
                    //Get the property off of the .. method ?? Most likely never happen.
                    var propertyValue = ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(_PropertyName,
                        BindingFlags.InvokeMethod, null, ActiveXObjectInstance.ActiveXObject, null);
                    return new ActiveXInstance(this.Prototype, propertyValue);
                }
                catch (Exception ex2)
                {
                    throw new JavaScriptException(Engine, "COMError", ex2.ToString());
                }
            }
             */
        }

        protected override void SetMissingPropertyValue(string propertyName, object value, bool throwOnError)
        {
            JuraScriptLibrary.ComReflector.ComReflectedMemberTypes memberType = ComReflector.GetMemberType(ActiveXObjectInstance.ActiveXObject, propertyName);
            if (memberType == ComReflector.ComReflectedMemberTypes.Field || memberType == ComReflector.ComReflectedMemberTypes.Property)
            {
                if (memberType == ComReflector.ComReflectedMemberTypes.Property)
                {
                    ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(propertyName, BindingFlags.SetProperty, null, ActiveXObjectInstance.ActiveXObject, new object[] { value });
                }
                else if (memberType == ComReflector.ComReflectedMemberTypes.Field)
                {
                    ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(propertyName, BindingFlags.SetField, null, ActiveXObjectInstance.ActiveXObject, new object[] { value });
                }
            }
            else
            {
                base.SetMissingPropertyValue(propertyName, value, throwOnError);
            }
            

            /* OLD METHOD
            try
            {
                ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(propertyName,
                    BindingFlags.SetProperty, null, ActiveXObjectInstance.ActiveXObject, new object[] { value });
            }
            catch (Exception ex)
            {
                base.SetMissingPropertyValue(propertyName, value, throwOnError);
            }
             */
        }

        /// <summary>
        /// Returns a string representing this object.
        /// </summary>
        /// <returns> A string representing this object. </returns>
        [JSFunction(Name = "toString")]
        public string ToStringJS()
        {
            JuraScriptLibrary.ComReflector.ComReflectedMemberTypes memberType = ComReflector.GetMemberType(ActiveXObjectInstance.ActiveXObject, _PropertyName);
            if(memberType == ComReflector.ComReflectedMemberTypes.Property) {
                var propertyValue = ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(_PropertyName, BindingFlags.GetProperty, null, ActiveXObjectInstance.ActiveXObject, null);
                return propertyValue.ToString();
            } else if(memberType == ComReflector.ComReflectedMemberTypes.Field) {
                var propertyValue = ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(_PropertyName, BindingFlags.GetProperty, null, ActiveXObjectInstance.ActiveXObject, null);
                return propertyValue.ToString();
            } else if(memberType == ComReflector.ComReflectedMemberTypes.Method) {
                var propertyValue = ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(_PropertyName,
                        BindingFlags.InvokeMethod, null, ActiveXObjectInstance.ActiveXObject, null);
                    return propertyValue.ToString();
            }
            throw new JavaScriptException(Engine, "COMError", "Error evaluating: " + _PropertyName);
            /* OLD METHOD
             try
            {
                var propertyValue = ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(_PropertyName,
                    BindingFlags.GetProperty, null, ActiveXObjectInstance.ActiveXObject, null);
                return propertyValue.ToString();
            }
            catch (Exception ex)
            {
                try
                {
                    //Get the property off of the .. method ?? Most likely never happen.
                    var propertyValue = ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(_PropertyName,
                        BindingFlags.InvokeMethod, null, ActiveXObjectInstance.ActiveXObject, null);
                    return propertyValue.ToString();
                }
                catch (Exception ex2)
                {
                    throw new JavaScriptException(Engine, "COMError", ex2.ToString());
                }
            }
             * */
        }


    }

    public class ActiveXInstance : ObjectInstance
    {
        private string ProgID = "";
        private string _ObjectHint = "";
        private object _ActiveXObject = null;
        public object ActiveXObject
        {
            get
            {
                return _ActiveXObject;
            }
            set
            {
                _ActiveXObject = value;
            }
           
        }
        private Type _ActiveXType = null;
        public ActiveXInstance(ObjectInstance prototype)
            : base(prototype)
        {
        }

        public ActiveXInstance(ObjectInstance prototype, string progID)
            : base(prototype)
        {
            ProgID = progID;

            _ActiveXType = System.Type.GetTypeFromProgID(progID);
            _ActiveXObject = Activator.CreateInstance(_ActiveXType);

            this.PopulateFunctions();
        }

        public ActiveXInstance(ObjectInstance prototype, object comobject)
            : base(prototype)
        {
            _ActiveXObject = comobject;
            this.PopulateFunctions();
        }
        public ActiveXInstance(ObjectInstance prototype, object comobject, string objecthint)
            : base(prototype)
        {
            _ActiveXObject = comobject;
            this.PopulateFunctions();
            _ObjectHint = objecthint;
        }

        public string ActiveXToString()
        {
            return _ActiveXObject.ToString();
        }

        protected override object GetMissingPropertyValue(string propertyName)
      {            
            try
            {
                //if (propertyName == "Add")
                //{ return new MyFunction(this.Engine, propertyName) { ActiveXObject = _ActiveXObject }; }
                ////Right now this is the only way I know of to check for a property
                ////Reflection does not see COM properties (ie: .GetType().GetMember())

                return new ActiveXFunction(this.Prototype, this, propertyName);

               // var propertyValue = new ActiveXInstance(this.Prototype, _ActiveXObject.GetType().InvokeMember(propertyName,
               //     BindingFlags.GetProperty, null, _ActiveXObject, null));
               // return propertyValue;
            }
            catch (Exception ex)
            {
                return base.GetMissingPropertyValue(propertyName);
            }
            
        }

        protected override void SetMissingPropertyValue(string propertyName, object value, bool throwOnError)
        {
            JuraScriptLibrary.ComReflector.ComReflectedMemberTypes memberType = ComReflector.GetMemberType(_ActiveXObject, propertyName);
            if (memberType == ComReflector.ComReflectedMemberTypes.Field || memberType == ComReflector.ComReflectedMemberTypes.Property)
            {
                if (memberType == ComReflector.ComReflectedMemberTypes.Property)
                {
                    _ActiveXObject.GetType().InvokeMember(propertyName, BindingFlags.SetProperty, null, _ActiveXObject, new object[] { value });
                }
                else if (memberType == ComReflector.ComReflectedMemberTypes.Field)
                {
                    _ActiveXObject.GetType().InvokeMember(propertyName, BindingFlags.SetField, null, _ActiveXObject, new object[] { value });
                }
            }
            else
            {
                base.SetMissingPropertyValue(propertyName, value, throwOnError);
            }
        }

        /* OLD METHOD
        protected override void SetMissingPropertyValue(string propertyName, object value, bool throwOnError)
        {
            try
            {
                _ActiveXObject.GetType().InvokeMember(propertyName,
                    BindingFlags.SetProperty, null, _ActiveXObject, new object[] { value });
            }
            catch (Exception ex)
            {
                base.SetMissingPropertyValue(propertyName, value, throwOnError);
            }
        }
         */

        //private class MyFunction : FunctionInstance
        //{
        //    private string name;
        //    public object ActiveXObject { get; set; }
        //    public MyFunction(ScriptEngine engine, string name)
        //        : base(engine)
        //    {
        //        this.name = name;
        //    }

        //    public override object CallLateBound(object thisObject, params object[] argumentValues)
        //    {
        //        if ((thisObject is ActiveXInstance) == false)
        //            throw new JavaScriptException(this.Engine, "TypeError", "Invalid 'this' value.");
        //        return ((ActiveXInstance)thisObject).CallFunction(this.name, argumentValues);
        //    }

        //    public override string ToString()
        //    {
        //        if (ActiveXObject != null)
        //        { return ActiveXObject.GetType().InvokeMember(name, BindingFlags.GetProperty, null, ActiveXObject, null).ToString(); }
        //        else { return base.ToString(); }
        //    }
        //}

        private object CallFunction(string functionName, params object[] argumentValues)
        {
            //return functionName + argumentValues.Length;
            //return "somepropertyok";
            /* OLD METHOD
            try
            {
                return _ActiveXObject.GetType().InvokeMember(functionName,
                    BindingFlags.InvokeMethod, null, _ActiveXObject, argumentValues);
            }
            catch (Exception ex)
            {
                throw new JavaScriptException(this.Engine, "COMError", ex.ToString());
            }
            throw new JavaScriptException(this.Engine, "COMError", "No such method.");
            */


            //Is it a property or a method?
            JuraScriptLibrary.ComReflector.ComReflectedMemberTypes memberType = ComReflector.GetMemberType(_ActiveXObject, functionName);
            if (memberType == ComReflector.ComReflectedMemberTypes.Method)
            {
                var memberValue = _ActiveXObject.GetType().InvokeMember(functionName,
                                        BindingFlags.InvokeMethod, null, _ActiveXObject, argumentValues);
                if (memberValue == null) { return null; }
                return new ActiveXInstance(this.Prototype, memberValue);
            }
            else if (memberType != ComReflector.ComReflectedMemberTypes.NOT_FOUND && memberType != ComReflector.ComReflectedMemberTypes.NO_ITYPEINFO)
            {
                //Sometimes CallFunction is getting a property
                //ie: Excel .Cells(x,x) which is an array but is called like a function even in JScript
                var memberValue = _ActiveXObject.GetType().InvokeMember(functionName,
                            BindingFlags.GetProperty, null, _ActiveXObject, null);
                if (memberValue == null) { return null; }
                return new ActiveXInstance(this.Prototype,
                        memberValue.GetType().InvokeMember(functionName, BindingFlags.GetProperty, null, memberValue, argumentValues));
            }

            throw new JavaScriptException(this.Engine, "COMError", "No such method.");
        }

        /// <summary>
        /// Returns a string representing this object.
        /// </summary>
        /// <returns> A string representing this object. </returns>
        [JSFunction(Name = "toString")]
        public string ToStringJS()
        {
            //Does it implement IDispatch ?
            var idisp = _ActiveXObject as IDispatch;

            if (idisp == null)
            {
                return _ActiveXObject.ToString();
            }
            else
            {
                //Does it have a default property? (dispid 0)
                ComReflector.ComReflectedMemberTypes memberType = ComReflector.ComReflectedMemberTypes.NOT_FOUND;
                string memberName = ComReflector.DefaultPropertyName(idisp, out memberType);
                if (memberName != "")
                {
                    return _ActiveXObject.GetType().InvokeMember(memberName, BindingFlags.GetProperty, null, _ActiveXObject, null).ToString();
                }
                else
                {
                    return _ActiveXObject.ToString();
                }
            }

            //OLD METHOD
            ////return this.ToString();
            //if (String.IsNullOrEmpty(_ObjectHint))
            //{
            //    return _ActiveXObject.ToString();
            //}
            //else
            //{
            //    if (_ObjectHint.ToUpper() == "DRIVES")
            //    {
            //        return _ActiveXObject.GetType().InvokeMember("Path", BindingFlags.GetProperty, null, _ActiveXObject, null).ToString();
            //    }
            //    return _ActiveXObject.ToString();
            //}
            
        }

    }

}
