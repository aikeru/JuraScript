using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using JuraScriptLibrary;
using System.IO;

namespace JuraScript.Tests
{
    /// <summary>
    /// Summary description for RequireTest
    /// </summary>
    [TestClass]
    public class RequireTest
    {
        public JuraScriptObject jso = new JuraScriptObject();

        public RequireTest()
        {
            File.WriteAllText("mymodule.jur", "exports.sayhello = function() { return \"hello from mymodule\"; };");
            File.WriteAllText("mymodule2.jur", "exports.sayhello = function() { return \"hello from other module\"; };");
            File.WriteAllText("mymodule3.jur", "exports.someProperty = \"property1\"; var myGlobal = \"globalvalue\"; exports.getGlobal = function() { return myGlobal; };");
            File.WriteAllText("mymodule4.jur", "exports = function() { return \"output\"; }");

        }


        [TestMethod]
        public void Test_Require()
        {
            jso.Execute(@"
                var myGlobal = ""private value"";
                var modResult1 = require(""mymodule"").sayhello();
                var module2 = require(""mymodule2.jur"");
                var modResult2 = module2.sayhello();
                var module3 = require(""mymodule3.jur"");
                var modResultGlobal = module3.getGlobal();
                var modResult3 = module3.someProperty;
                require(""mymodule3.jur"");
            ");

            Assert.AreEqual("private value", jso.JurassicEngine.GetGlobalValue("myGlobal"));
            Assert.AreEqual("hello from mymodule", jso.JurassicEngine.GetGlobalValue("modResult1"));
            Assert.AreEqual("hello from other module", jso.JurassicEngine.GetGlobalValue("modResult2"));
            Assert.AreEqual("globalvalue", jso.JurassicEngine.GetGlobalValue("modResultGlobal"));
            Assert.AreEqual("property1", jso.JurassicEngine.GetGlobalValue("modResult3"));

            //Test assigning value to the exports object itself ...
            jso.Execute(@"var modResult4 = require(""mymodule4.jur"")();");

            Assert.AreEqual("output", jso.JurassicEngine.GetGlobalValue("modResult4"));
            
            //Test loading a built-in module
            jso.Execute(@"var modResult5 = require(""assert"");");
            
        }
    }
}
