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
            Debug.WriteLine("constructor: ActiveXFunction(prototype, " + ComReflector.GetCOMObjectTypeName(instanceobject) + ", " + propertyName + ")");
            ActiveXObjectInstance = instanceobject;
            _PropertyName = propertyName;
            this.PopulateFunctions();
        }

        public object ActiveXUnpack()
        {
            Debug.WriteLine("ActiveXFunction.ActiveXUnpack() property name is " + _PropertyName);
            ComReflector.ComReflectedMemberTypes memberType = ComReflector.GetMemberType(ActiveXObjectInstance.ActiveXObject, _PropertyName);
            Debug.WriteLine("ActiveXFunction.ActiveXUnpack - got memberType: " + memberType.ToString());
            if (memberType == ComReflector.ComReflectedMemberTypes.Field)
            {
                Debug.WriteLine("ActiveXFunction.ActiveXUnpack - returning field value [" + ComReflector.GetCOMObjectTypeName(ActiveXObjectInstance.ActiveXObject) + "]");
                return ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(_PropertyName,
                    BindingFlags.GetField, null, ActiveXObjectInstance.ActiveXObject, null);
            }
            else if (memberType == ComReflector.ComReflectedMemberTypes.Property)
            {
                Debug.WriteLine("ActiveXFunction.ActiveXUnpack - returning Property value [" + ComReflector.GetCOMObjectTypeName(ActiveXObjectInstance.ActiveXObject) + "]");
                return ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(_PropertyName,
                    BindingFlags.GetProperty, null, ActiveXObjectInstance.ActiveXObject, null);
            }
            else if (memberType == ComReflector.ComReflectedMemberTypes.Method)
            {
                Debug.WriteLine("ActiveXFunction.ActiveXUnpack - returning Method value [" + ComReflector.GetCOMObjectTypeName(ActiveXObjectInstance.ActiveXObject) + "]");
                return ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(_PropertyName,
                    BindingFlags.InvokeMethod, null, ActiveXObjectInstance.ActiveXObject, null);
            }
            throw new JavaScriptException(Engine, "COMError", "Error unpacking: " + _PropertyName);
        }

        private object ActiveXGetValue(string propertyName)
        {
            Debug.WriteLine("ActiveXFunction.ActiveXGetValue(" + propertyName +  ")");
            JuraScriptLibrary.ComReflector.ComReflectedMemberTypes memberType = ComReflector.GetMemberType(ActiveXObjectInstance.ActiveXObject, _PropertyName);
            Debug.WriteLine("ActiveXFunction.ActiveXGetValue - got memberType: " + memberType.ToString());
            if (memberType == ComReflector.ComReflectedMemberTypes.Method)
            {
                Debug.WriteLine("ActiveXFunction.ActiveXGetValue - invoking method");
                //Get the property off of the .. method ?? Most likely never happen.
                var propertyValue = new ActiveXInstance(this.Prototype, ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(_PropertyName,
                        BindingFlags.InvokeMethod, null, ActiveXObjectInstance.ActiveXObject, null));
                Debug.WriteLine("ActiveXFunction.ActiveXGetValue - wrapping propertyValue [" + ComReflector.GetCOMObjectTypeName(propertyValue) + "] with ActiveXFunction");
                return new ActiveXFunction(this.Prototype, propertyValue, propertyName);
            }
            else if (memberType != ComReflector.ComReflectedMemberTypes.NOT_FOUND && memberType != ComReflector.ComReflectedMemberTypes.NO_ITYPEINFO)
            {
                Debug.WriteLine("ActiveXFunction.ActiveXGetValue - invoking property");
                //Unwrap the ActiveXInstance ActiveX Object, then create a new ActiveXFunction pointing at the new property
                var propertyValue = new ActiveXInstance(this.Prototype, ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(_PropertyName,
                    BindingFlags.GetProperty, null, ActiveXObjectInstance.ActiveXObject, null));

                Debug.WriteLine("ActiveXFunction.ActiveXGetValue - wrapping propertyValue [" + ComReflector.GetCOMObjectTypeName(propertyValue) + "] with ActiveXFunction");

                return new ActiveXFunction(this.Prototype, propertyValue, propertyName);
            }

            throw new JavaScriptException(this.Engine, "COMError", "No such property.");
        }

        protected override object GetMissingPropertyValue(string propertyName)
        {
            Debug.WriteLine("ActiveXFunction.GetMissingPropertyValue(" + propertyName + ") calling ActiveXGetValue(" + propertyName + ")");
            return ActiveXGetValue(propertyName);
        }

        public override object CallLateBound(object thisObject, params object[] argumentValues)
        {
            //if ((thisObject is ActiveXFunction) == false)
            //    throw new JavaScriptException(this.Engine, "TypeError", "Invalid 'this' value.");
            //return ((ActiveXFunction)thisObject).CallFunction(this._PropertyName, argumentValues);
            Debug.WriteLine("ActiveXFunction.CallLateBound() passing to CallFunction()");

            return CallFunction(this._PropertyName, argumentValues);
        }

        private object CallFunction(string functionName, params object[] argumentValues)
        {
            string arguments = String.Join(",", argumentValues.Select(a => a != null ? a.ToString() : "(null)").ToArray());
            Debug.WriteLine(String.Format("ActiveXFunction.CallFunction({0}, {1})", functionName, arguments));

            //Is it a property or a method?
            JuraScriptLibrary.ComReflector.ComReflectedMemberTypes memberType = ComReflector.GetMemberType(ActiveXObjectInstance.ActiveXObject, _PropertyName);
            Debug.WriteLine("ActiveXFunction.CallFunction - got memberType: " + memberType.ToString());

            if(memberType == ComReflector.ComReflectedMemberTypes.Method) {
                Debug.WriteLine("ActiveXFunction.CallFunction - invoking method");
                var memberValue = ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(_PropertyName,
                                        BindingFlags.InvokeMethod, null, ActiveXObjectInstance.ActiveXObject, argumentValues);
                if(memberValue == null) {
                    Debug.WriteLine("ActiveXFunction.CallFunction - memberValue returned null, returning null (is this an error? no ref?)");
                    return null; 
                }
                Debug.WriteLine("ActiveXFunction.CallFunction - returning wrapped ActiveXInstance of " + ComReflector.GetCOMObjectTypeName(memberValue));
                return new ActiveXInstance(this.Prototype, memberValue);
            } else if(memberType != ComReflector.ComReflectedMemberTypes.NOT_FOUND && memberType != ComReflector.ComReflectedMemberTypes.NO_ITYPEINFO) {
                //Sometimes CallFunction is getting a property
                //ie: Excel .Cells(x,x) which is an array but is called like a function even in JScript
                
                //var memberValue = ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(_PropertyName,
                //            BindingFlags.GetProperty, null, ActiveXObjectInstance.ActiveXObject, null);

                //if (memberValue == null) { return null; }
                //return new ActiveXInstance(this.Prototype,
                //        memberValue.GetType().InvokeMember(functionName, BindingFlags.GetProperty, null, memberValue, argumentValues));
                Debug.WriteLine("ActiveXFunction.CallFunction - invoking property on " + ComReflector.GetCOMObjectTypeName(ActiveXObjectInstance.ActiveXObject));

                var member = ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(_PropertyName, BindingFlags.GetProperty, null, ActiveXObjectInstance.ActiveXObject, null);

                Debug.WriteLine("ActiveXFunction.CallFunction - member type name: " + ComReflector.GetCOMObjectTypeName(member));

                Debug.WriteLine("ActiveXFunction.CallFunction - looking for a default member...");
                //Find the default member and invoke it here
                string defaultName = ComReflector.DefaultPropertyName(member as IDispatch, out memberType);
                
                if (!String.IsNullOrWhiteSpace(defaultName))
                {
                    Debug.WriteLine("ActiveXFunction.CallFunction - found member: " + defaultName + " returning invoke GetProperty on member ["  + ComReflector.GetCOMObjectTypeName(member) +  "] with argumentValues [" + arguments + "]");
                    return new ActiveXInstance(this.Prototype,
                        member.GetType().InvokeMember(defaultName, BindingFlags.GetProperty, null, member, argumentValues));
                }
                //No default member... I guess it's the object itself
                else
                {
                    Debug.WriteLine("ActiveXFunction.CallFunction - no default member was found, wrapping the object itself [" + ComReflector.GetCOMObjectTypeName(ActiveXObjectInstance.ActiveXObject) + "] with argumentValues [" + arguments + "]");
                    return new ActiveXInstance(this.Prototype,
                            ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(functionName, BindingFlags.GetProperty, null, ActiveXObjectInstance.ActiveXObject, argumentValues));
                }
            }
            Debug.WriteLine("ActiveXFunction.CallFunction - throwing COMError no such method");
            throw new JavaScriptException(this.Engine, "COMError", "No such method.");
        }

        /// <summary>
        /// Returns a primitive value associated with the object.
        /// </summary>
        /// <returns> A primitive value associated with the object. </returns>
        [JSFunction(Name = "valueOf")]
        public new ObjectInstance ValueOf()
        {
            Debug.WriteLine("ActiveXFunction.ValueOf()");
            JuraScriptLibrary.ComReflector.ComReflectedMemberTypes memberType = ComReflector.GetMemberType(ActiveXObjectInstance.ActiveXObject, _PropertyName);
            Debug.WriteLine("ActiveXFunction.ValueOf - got memberType:" + memberType.ToString());
            if (memberType == ComReflector.ComReflectedMemberTypes.Method)
            {
                Debug.WriteLine("ActiveXFunction.ValueOf - invoking method on " + ComReflector.GetCOMObjectTypeName(ActiveXObjectInstance.ActiveXObject));
                //Get the property off of the .. method ?? Most likely never happen.
                var propertyValue = new ActiveXInstance(this.Prototype, ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(_PropertyName,
                        BindingFlags.InvokeMethod, null, ActiveXObjectInstance.ActiveXObject, null));

                Debug.WriteLine("ActiveXFunction.ValueOf - returning wrapped ActiveXInstance " + ComReflector.GetCOMObjectTypeName(propertyValue));

                return new ActiveXInstance(this.Prototype, propertyValue);
            }
            else if (memberType != ComReflector.ComReflectedMemberTypes.NOT_FOUND && memberType != ComReflector.ComReflectedMemberTypes.NO_ITYPEINFO)
            {
                Debug.WriteLine("ActiveXFunction.ValueOf - invoking property on " + ComReflector.GetCOMObjectTypeName(ActiveXObjectInstance.ActiveXObject));
                //Unwrap the ActiveXInstance ActiveX Object, then create a new ActiveXFunction pointing at the new property
                var propertyValue = new ActiveXInstance(this.Prototype, ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(_PropertyName,
                    BindingFlags.GetProperty, null, ActiveXObjectInstance.ActiveXObject, null));

                Debug.WriteLine("ActiveXFunction.ValueOf - returning wrapped ActiveXInstance " + ComReflector.GetCOMObjectTypeName(propertyValue));

                return new ActiveXInstance(this.Prototype, propertyValue);
            }

            throw new JavaScriptException(this.Engine, "COMError", "Error evaluating: " + _PropertyName);
        }

        protected override void SetMissingPropertyValue(string propertyName, object value, bool throwOnError)
        {
            Debug.WriteLine(String.Format("ActiveXFunction.SetMissingPropertyValue({0}, {1}, throwOnError)", propertyName, value ?? "null"));
            //JuraScriptLibrary.ComReflector.ComReflectedMemberTypes memberType = ComReflector.GetMemberType(ActiveXObjectInstance.ActiveXObject, propertyName);
            //Get the member type of _PropertyName (the current property/object)

            object memberValue = ComReflector.InvokeMember(ActiveXObjectInstance.ActiveXObject, _PropertyName, null, false);
            ComReflector.InvokeMember(memberValue, propertyName, new object[] { value } , true);

            JuraScriptLibrary.ComReflector.ComReflectedMemberTypes memberType = ComReflector.GetMemberType(ActiveXObjectInstance.ActiveXObject, _PropertyName);
            Debug.WriteLine("ActiveXFunction.SetMissingPropertyValue - got memberType: " + memberType.ToString());
            
        }

        /// <summary>
        /// Returns a string representing this object.
        /// </summary>
        /// <returns> A string representing this object. </returns>
        [JSFunction(Name = "toString")]
        public string ToStringJS()
        {
            Debug.WriteLine("ActiveXFunction.ToStringJS()");
            JuraScriptLibrary.ComReflector.ComReflectedMemberTypes memberType = ComReflector.GetMemberType(ActiveXObjectInstance.ActiveXObject, _PropertyName);
            Debug.WriteLine("ActiveXFunction.ToStringJS - got memberType: " + memberType.ToString() + " of _PropertyName: " + _PropertyName);
            if(memberType == ComReflector.ComReflectedMemberTypes.Property) {
                Debug.WriteLine("ActiveXFunction.ToStringJS - invoking/returning GetProperty.ToString()");
                var propertyValue = ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(_PropertyName, BindingFlags.GetProperty, null, ActiveXObjectInstance.ActiveXObject, null);
                return propertyValue.ToString();
            } else if(memberType == ComReflector.ComReflectedMemberTypes.Field) {
                Debug.WriteLine("ActiveXFunction.ToStringJS - invoking/returning GetField.ToString() (will probably fail)");
                var propertyValue = ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(_PropertyName, BindingFlags.GetProperty, null, ActiveXObjectInstance.ActiveXObject, null);
                return propertyValue.ToString();
            } else if(memberType == ComReflector.ComReflectedMemberTypes.Method) {
                Debug.WriteLine("ActiveXFunction.ToStringJS - invoking/returning Method with null args. The result.ToString() will be returned.");
                var propertyValue = ActiveXObjectInstance.ActiveXObject.GetType().InvokeMember(_PropertyName,
                        BindingFlags.InvokeMethod, null, ActiveXObjectInstance.ActiveXObject, null);
                    return propertyValue.ToString();
            }
            Debug.WriteLine("ActiveXFunction.ToStringJS - throwing a COMError");
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
            Debug.WriteLine("constructor: ActiveXInstance.ActiveXInstance(prototype)");
        }

        public ActiveXInstance(ObjectInstance prototype, string progID)
            : base(prototype)
        {
            Debug.WriteLine("constructor: ActiveXInstance.ActiveXInstance(prototype, " + progID + ")");
            ProgID = progID;

            _ActiveXType = System.Type.GetTypeFromProgID(progID);
            _ActiveXObject = Activator.CreateInstance(_ActiveXType);

            this.PopulateFunctions();
        }

        public ActiveXInstance(ObjectInstance prototype, object comobject)
            : base(prototype)
        {
            Debug.WriteLine("constructor: ActiveXInstance.ActiveXInstance(prototype, " + ComReflector.GetCOMObjectTypeName(comobject) + ")");
            
            _ActiveXObject = comobject;
            this.PopulateFunctions();
        }

        protected override object GetMissingPropertyValue(string propertyName)
        {
            Debug.WriteLine("ActiveXInstance.GetMissingPropertyValue(" + propertyName + ") -- returning a new ActiveXFunction object");
            return new ActiveXFunction(this.Prototype, this, propertyName);
        }

        protected override void SetMissingPropertyValue(string propertyName, object value, bool throwOnError)
        {
            Debug.WriteLine("ActiveXInstance.SetMissingPropertyValue(" + propertyName + ", " + value.ToString() + ", throwOnError)");
            JuraScriptLibrary.ComReflector.ComReflectedMemberTypes memberType = ComReflector.GetMemberType(_ActiveXObject, propertyName);
            Debug.WriteLine("ActiveXInstance.SetMissingPropertyValue - got memberType: " + memberType.ToString());
            if (memberType == ComReflector.ComReflectedMemberTypes.Field || memberType == ComReflector.ComReflectedMemberTypes.Property)
            {
                if (memberType == ComReflector.ComReflectedMemberTypes.Property)
                {
                    Debug.WriteLine("ActiveXInstance.SetMissingPropertyValue - setting property to value");
                    _ActiveXObject.GetType().InvokeMember(propertyName, BindingFlags.SetProperty, null, _ActiveXObject, new object[] { value });
                }
                else if (memberType == ComReflector.ComReflectedMemberTypes.Field)
                {
                    Debug.WriteLine("ActiveXInstance.SetMissingPropertyValue - setting field to value");
                    _ActiveXObject.GetType().InvokeMember(propertyName, BindingFlags.SetField, null, _ActiveXObject, new object[] { value });
                }
            }
            else
            {
                Debug.WriteLine("ActiveXInstance.SetMissingPropertyValue - didn't find it, calling base...");
                base.SetMissingPropertyValue(propertyName, value, throwOnError);
            }
        }

        

        private object CallFunction(string functionName, params object[] argumentValues)
        {
            string arguments = String.Join(",", argumentValues.Select(a => a != null ? a.ToString() : "(null)").ToArray());
            Debug.WriteLine("ActiveXInstance.CallFunction(" + functionName + ", " + arguments + ") ");
            //Is it a property or a method?
            JuraScriptLibrary.ComReflector.ComReflectedMemberTypes memberType = ComReflector.GetMemberType(_ActiveXObject, functionName);
            Debug.WriteLine("ActiveXInstance.CallFunction - got memberType: " + memberType.ToString());
            if (memberType == ComReflector.ComReflectedMemberTypes.Method)
            {
                Debug.WriteLine("ActiveXInstance.CallFunction - invoking method " + functionName);
                var memberValue = _ActiveXObject.GetType().InvokeMember(functionName,
                                        BindingFlags.InvokeMethod, null, _ActiveXObject, argumentValues);

                if (memberValue == null) {
                    Debug.WriteLine("ActiveXInstance.CallFunction - returned a null membervalue. Returning null.");
                    return null; 
                }
                Debug.WriteLine("ActiveXInstance.CallFunction - returning a new ActiveXInstance wrapping memberValue: " + memberValue.ToString());
                return new ActiveXInstance(this.Prototype, memberValue);
            }
            else if (memberType != ComReflector.ComReflectedMemberTypes.NOT_FOUND && memberType != ComReflector.ComReflectedMemberTypes.NO_ITYPEINFO)
            {
                //Sometimes CallFunction is getting a property
                //ie: Excel .Cells(x,x) which is an array but is called like a function even in JScript
                Debug.WriteLine("ActiveXInstance.CallFunction - invoking a property");

                var memberValue = _ActiveXObject.GetType().InvokeMember(functionName,
                            BindingFlags.GetProperty, null, _ActiveXObject, null);
                if (memberValue == null) {
                    Debug.WriteLine("ActiveXInstance.CallFunction - returned a null GetProperty value. Returning null. Is this an error - no ref?");
                    return null; 
                }
                Debug.WriteLine("ActiveXInstance.CallFunction - returning a new ActiveXInstance wrapping memberValue: " + memberValue.ToString());
                return new ActiveXInstance(this.Prototype,
                        memberValue.GetType().InvokeMember(functionName, BindingFlags.GetProperty, null, memberValue, argumentValues));
            }
            Debug.WriteLine("ActiveXInstance.CallFunction - throwing a COMError, didn't find it.");
            throw new JavaScriptException(this.Engine, "COMError", "No such method.");
        }

        /// <summary>
        /// Returns a string representing this object.
        /// </summary>
        /// <returns> A string representing this object. </returns>
        [JSFunction(Name = "toString")]
        public string ToStringJS()
        {
            Debug.WriteLine("ActiveXInstance.ToStringJS()");
            //Does it implement IDispatch ?
            var idisp = _ActiveXObject as IDispatch;

            if (idisp == null)
            {
                Debug.WriteLine("ActiveXInstance.ToStringJS - doesn't implement IDispatch, so returning .ToString()");
                return _ActiveXObject.ToString();
            }
            else
            {
                Debug.WriteLine("ActiveXInstance.ToStringJS - IDispatch OK, looking for default property ...");
                //Does it have a default property? (dispid 0)
                ComReflector.ComReflectedMemberTypes memberType = ComReflector.ComReflectedMemberTypes.NOT_FOUND;
                string memberName = ComReflector.DefaultPropertyName(idisp, out memberType);
                if (memberName != "")
                {
                    Debug.WriteLine("ActiveXInstance.ToStringJS - found member " + memberName + " of type " + memberType.ToString() + " invoking this and returning raw value.");
                    return _ActiveXObject.GetType().InvokeMember(memberName, BindingFlags.GetProperty, null, _ActiveXObject, null).ToString();
                }
                else
                {
                    Debug.WriteLine("ActiveXInstance.ToStringJS - didn't find anything that looked like default. Returning .ToString()");
                    return _ActiveXObject.ToString();
                }
            }
            
        }

    }

}
