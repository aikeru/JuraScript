using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Jurassic.Library;
using Jurassic;

namespace JurassicTest
{
    public class ActiveXConstructor : ClrFunction
    {

        public ActiveXConstructor(ScriptEngine engine)
            : base(engine.Function.InstancePrototype, "ActiveXObject", new ActiveXInstance(engine.Object.InstancePrototype))
        {
        }

        [JSConstructorFunction]
        public ActiveXInstance Construct()
        {
            return new ActiveXInstance(this.Prototype);
        }

        [JSConstructorFunction]
        public ActiveXInstance Construct(string progID)
        {
            return new ActiveXInstance(this.Prototype, progID);
        }

    }
}
