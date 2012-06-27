using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Jurassic.Library;
using Jurassic;

namespace JuraScriptLibrary.Require
{
    public class exportsObject : ObjectInstance
    {
        public Dictionary<string, object> ExportProperties = new Dictionary<string, object>();

        public exportsObject(ObjectInstance prototype)
            : base(prototype)
        {
            this.PopulateFunctions();
        }

        protected override object GetMissingPropertyValue(string propertyName)
        {
            if (ExportProperties.ContainsKey(propertyName))
            {
                return ExportProperties[propertyName];
            }
            else
            {
                return Undefined.Value;
            }
        }
        protected override void SetMissingPropertyValue(string propertyName, object value, bool throwOnError)
        {
            if (ExportProperties.ContainsKey(propertyName))
            {
                ExportProperties[propertyName] = value;
            }
            else
            {
                ExportProperties.Add(propertyName, value);
            }
        }
    }
}
