using JuraScriptLibrary.Require;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using Jurassic.Library;

namespace JuraScript.CommonJS.Tests
{
    
    
    /// <summary>
    ///This is a test class for requireObjectTest and is intended
    ///to contain all requireObjectTest Unit Tests
    ///</summary>
    [TestClass()]
    public class requireObjectTest
    {


        private TestContext testContextInstance;

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        #region Additional test attributes
        // 
        //You can use the following additional attributes as you write your tests:
        //
        //Use ClassInitialize to run code before running the first test in the class
        //[ClassInitialize()]
        //public static void MyClassInitialize(TestContext testContext)
        //{
        //}
        //
        //Use ClassCleanup to run code after all tests in a class have run
        //[ClassCleanup()]
        //public static void MyClassCleanup()
        //{
        //}
        //
        //Use TestInitialize to run code before running each test
        //[TestInitialize()]
        //public void MyTestInitialize()
        //{
        //}
        //
        //Use TestCleanup to run code after each test has run
        //[TestCleanup()]
        //public void MyTestCleanup()
        //{
        //}
        //
        #endregion


        /// <summary>
        ///A test for requireObject Constructor
        ///</summary>
        [TestMethod()]
        public void requireObjectConstructorTest()
        {
            ObjectInstance prototype = null; // TODO: Initialize to an appropriate value
            requireObject target = new requireObject(prototype);
            Assert.Inconclusive("TODO: Implement code to verify target");
        }

        /// <summary>
        ///A test for CallLateBound
        ///</summary>
        [TestMethod()]
        public void CallLateBoundTest()
        {
            ObjectInstance prototype = null; // TODO: Initialize to an appropriate value
            requireObject target = new requireObject(prototype); // TODO: Initialize to an appropriate value
            object thisObject = null; // TODO: Initialize to an appropriate value
            object[] argumentValues = null; // TODO: Initialize to an appropriate value
            object expected = null; // TODO: Initialize to an appropriate value
            object actual;
            actual = target.CallLateBound(thisObject, argumentValues);
            Assert.AreEqual(expected, actual);
            Assert.Inconclusive("Verify the correctness of this test method.");
        }
    }
}
