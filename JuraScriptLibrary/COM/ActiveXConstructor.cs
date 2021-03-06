﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Jurassic.Library;
using Jurassic;
using System.Diagnostics;

namespace JuraScriptLibrary.COM
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
            throw new JavaScriptException(Engine, "COMError", "Must specify a COM type to instantiate.");
        }

        [JSConstructorFunction]
        public ActiveXInstance Construct(string progID)
        {
            return new ActiveXInstance(this.Prototype, progID);
        }

    }
}
