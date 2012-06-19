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
using System.Diagnostics;

namespace JuraScriptLibrary.COM
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
            ComReflector.ComReflectedMemberTypes memberType = ComReflector.GetMemberType(ActiveXObjectInstance.ActiveXObject, _PropertyName);
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
            ComReflector.ComReflectedMemberTypes memberType = ComReflector.GetMemberType(ActiveXObjectInstance.ActiveXObject, _PropertyName);

            if (memberType == ComReflector.ComReflectedMemberTypes.Method)
            {
                var memberValue = ComReflector.InvokeMember(ActiveXObjectInstance.ActiveXObject, _PropertyName, argumentValues, false);
                if (memberValue == null) { return memberValue; }
                return new ActiveXInstance(this.Prototype, memberValue);
            }
            else
            {
                var memberValue = ComReflector.InvokeMember(ActiveXObjectInstance.ActiveXObject, _PropertyName, null, false);
                //Find the default member and invoke it here
                string defaultName = ComReflector.DefaultPropertyName(memberValue as IDispatch, out memberType);
                if (!String.IsNullOrWhiteSpace(defaultName))
                {
                    return new ActiveXInstance(this.Prototype, ComReflector.InvokeMember(memberValue, defaultName, argumentValues, false));
                }
                else if (memberType != ComReflector.ComReflectedMemberTypes.NOT_FOUND && memberType != ComReflector.ComReflectedMemberTypes.NO_ITYPEINFO) //No default member... I guess it's the object itself
                {
                    return new ActiveXInstance(this.Prototype, ComReflector.InvokeMember(ActiveXObjectInstance.ActiveXObject, functionName, argumentValues, false));
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
            ComReflector.ComReflectedMemberTypes memberType = ComReflector.GetMemberType(ActiveXObjectInstance.ActiveXObject, _PropertyName);
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
            //JuraScriptLibrary.ComReflector.ComReflectedMemberTypes memberType = ComReflector.GetMemberType(ActiveXObjectInstance.ActiveXObject, propertyName);
            //Get the member type of _PropertyName (the current property/object)

            object memberValue = ComReflector.InvokeMember(ActiveXObjectInstance.ActiveXObject, _PropertyName, null, false);
            ComReflector.InvokeMember(memberValue, propertyName, new object[] { value }, true);

            ComReflector.ComReflectedMemberTypes memberType = ComReflector.GetMemberType(ActiveXObjectInstance.ActiveXObject, _PropertyName);

        }

        /// <summary>
        /// Returns a string representing this object.
        /// </summary>
        /// <returns> A string representing this object. </returns>
        [JSFunction(Name = "toString")]
        public string ToStringJS()
        {
            ComReflector.ComReflectedMemberTypes memberType = ComReflector.GetMemberType(ActiveXObjectInstance.ActiveXObject, _PropertyName);

            return ComReflector.InvokeMember(ActiveXObjectInstance.ActiveXObject, _PropertyName, null, false).ToString();
        }


    }
}
