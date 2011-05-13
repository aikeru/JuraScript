using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Jurassic.Library;
using Jurassic;

namespace JurassicTest
{
    public class EnumeratorConstructor : ClrFunction
    {
        public EnumeratorConstructor(ScriptEngine engine)
            : base(engine.Function.InstancePrototype, "Enumerator", new ActiveXInstance(engine.Object.InstancePrototype))
        {
        }

        [JSConstructorFunction]
        public EnumeratorInstance Construct()
        {
            return new EnumeratorInstance(this.Prototype);
        }

        [JSConstructorFunction]
        public EnumeratorInstance Construct(object axObj)
        {
            return new EnumeratorInstance(this.Prototype, axObj);
        }
    }
}
