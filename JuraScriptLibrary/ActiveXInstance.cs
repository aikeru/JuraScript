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

            ComReflector.ComReflectedMemberTypes memberType = ComReflector.GetMemberType(ActiveXObjectInstance.ActiveXObject, _PropertyName);
            if (memberType == ComReflector.ComReflectedMemberTypes.Field)
            {
                return ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(_PropertyName,
                    BindingFlags.GetField, null, ActiveXObjectInstance.ActiveXObject, null);
            }
            else if (memberType == ComReflector.ComReflectedMemberTypes.Property)
            {
                return ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(_PropertyName,
                    BindingFlags.GetProperty, null, ActiveXObjectInstance.ActiveXObject, null);
            }
            else if (memberType == ComReflector.ComReflectedMemberTypes.Method)
            {
                return ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(_PropertyName,
                    BindingFlags.InvokeMethod, null, ActiveXObjectInstance.ActiveXObject, null);
            }
            throw new JavaScriptException(Engine, "COMError", "Error unpacking: " + _PropertyName);
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
                
                //var memberValue = ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(_PropertyName,
                //            BindingFlags.GetProperty, null, ActiveXObjectInstance.ActiveXObject, null);

                //if (memberValue == null) { return null; }
                //return new ActiveXInstance(this.Prototype,
                //        memberValue.GetType().InvokeMember(functionName, BindingFlags.GetProperty, null, memberValue, argumentValues));

                var member = ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(_PropertyName, BindingFlags.GetProperty, null, ActiveXObjectInstance.ActiveXObject, null);
                //Find the default member and invoke it here
                string defaultName = ComReflector.DefaultPropertyName(member as IDispatch, out memberType);
                if (!String.IsNullOrWhiteSpace(defaultName))
                {
                    return new ActiveXInstance(this.Prototype,
                        member.GetType().InvokeMember(defaultName, BindingFlags.GetProperty, null, member, argumentValues));
                }
                //No default member... I guess it's the object itself
                else
                {
                    return new ActiveXInstance(this.Prototype,
                            ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(functionName, BindingFlags.GetProperty, null, ActiveXObjectInstance.ActiveXObject, argumentValues));
                }
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
        }


    }

    public class ActiveXInstance : ObjectInstance
    {
        private string ProgID = "";
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

        public string ActiveXToString()
        {
            return _ActiveXObject.ToString();
        }

        protected override object GetMissingPropertyValue(string propertyName)
        {
            return new ActiveXFunction(this.Prototype, this, propertyName);
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

        

        private object CallFunction(string functionName, params object[] argumentValues)
        {
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
            
        }

    }

}
