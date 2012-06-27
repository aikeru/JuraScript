using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Jurassic.Library;

namespace JuraScriptLibrary.Require
{
    public class exportsObject : ObjectInstance
    {
        public exportsObject(ObjectInstance prototype)
            : base(prototype)
        {
            this.PopulateFunctions();
        }
    }
}
