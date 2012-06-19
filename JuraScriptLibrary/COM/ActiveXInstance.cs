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

        protected override object GetMissingPropertyValue(string propertyName)
        {
            return new ActiveXFunction(this.Prototype, this, propertyName);
        }

        protected override void SetMissingPropertyValue(string propertyName, object value, bool throwOnError)
        {
            ComReflector.ComReflectedMemberTypes memberType = ComReflector.GetMemberType(_ActiveXObject, propertyName);
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
            string arguments = String.Join(",", argumentValues.Select(a => a != null ? a.ToString() : "(null)").ToArray());
            //Is it a property or a method?
            ComReflector.ComReflectedMemberTypes memberType = ComReflector.GetMemberType(_ActiveXObject, functionName);
            if (memberType == ComReflector.ComReflectedMemberTypes.Method)
            {
                var memberValue = _ActiveXObject.GetType().InvokeMember(functionName,
                                        BindingFlags.InvokeMethod, null, _ActiveXObject, argumentValues);

                if (memberValue == null) {
                    return null; 
                }
                return new ActiveXInstance(this.Prototype, memberValue);
            }
            else if (memberType != ComReflector.ComReflectedMemberTypes.NOT_FOUND && memberType != ComReflector.ComReflectedMemberTypes.NO_ITYPEINFO)
            {
                //Sometimes CallFunction is getting a property
                //ie: Excel .Cells(x,x) which is an array but is called like a function even in JScript
                var memberValue = _ActiveXObject.GetType().InvokeMember(functionName,
                            BindingFlags.GetProperty, null, _ActiveXObject, null);
                if (memberValue == null) {
                    return null; 
                }
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
