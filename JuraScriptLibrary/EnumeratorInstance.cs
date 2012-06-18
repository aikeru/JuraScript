using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Jurassic.Library;
using System.Collections;

namespace JurassicTest
{

    public class EnumeratorInstance : ObjectInstance
    {
        /// <summary>
        /// WScript's Enumerator automatically grabs the first one
        /// </summary>
        private bool FirstUseBit = false;
        private object _ActiveXObject = null;
        private IEnumerable _Enumerable = null;
        private IEnumerator _Enumerator = null;
        private List<object> _Stack = new List<object>();

        private string ObjectHint = "";

        public EnumeratorInstance(ObjectInstance prototype)
            : base(prototype)
        {
        }
        public EnumeratorInstance(ObjectInstance prototype, object axObj)
            : base(prototype)
        {
            if (axObj is ActiveXFunction)
            {
                ActiveXFunction axFunc = axObj as ActiveXFunction;
                ObjectHint = axFunc.PropertyName;
                _ActiveXObject = axFunc.ActiveXUnpack();
            }
            else if (axObj is ActiveXInstance)
            {
                _ActiveXObject = ((ActiveXInstance)axObj).ActiveXObject;
            }
            else
            {
                _ActiveXObject = axObj;
            }
            _Enumerable = _ActiveXObject as IEnumerable;
            if (_Enumerable != null)
            {
                _Enumerator = _Enumerable.GetEnumerator();
            }
            this.PopulateFunctions();
        }
        [JSFunction]
        public void moveFirst()
        {
            if (_Enumerator != null)
            {
                _Enumerator.Reset();
            }
        }
        [JSFunction]
        public bool atEnd()
        {
            if (_Enumerator == null) { return true; }
            if (_Stack.Count > 0)
            {
                return false;
            }
            else
            {
                if (_Enumerator.Current != null)
                {
                    _Stack.Add(_Enumerator.Current);
                }
                if (_Enumerator.MoveNext())
                {
                    return false;
                }
            }
            return true;
        }
        [JSFunction]
        public object item()
        {
            if (_Stack.Count > 0)
            {
                return new ActiveXInstance(this.Prototype, _Stack[0]);
            }
            else
            {
                if (!FirstUseBit && _Enumerator.Current == null)
                {
                    FirstUseBit = true;
                    _Enumerator.MoveNext();
                }
                ActiveXInstance ax = new ActiveXInstance(this.Prototype, _Enumerator.Current);
                return ax;
            }
        }
        [JSFunction]
        public void moveNext()
        {
            if (_Stack.Count > 0)
            {
                _Stack.RemoveAt(0);
            }
            else
            {
                _Enumerator.MoveNext();
            }
        }
    }
}
