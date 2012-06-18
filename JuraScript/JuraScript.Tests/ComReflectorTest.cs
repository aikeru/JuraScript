using JuraScriptLibrary;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Runtime.InteropServices.ComTypes;
using System.IO;
using Jurassic.Library;

namespace JuraScript.Tests
{
    
    
    /// <summary>
    ///This is a test class for ComReflectorTest and is intended
    ///to contain all ComReflectorTest Unit Tests
    ///</summary>
    [TestClass()]
    public class ComReflectorTest
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

        [TestMethod()]
        public void COMTest_FileSystemObject_EnumerateDrives()
        {
            //Enumerate drives
            JuraScriptObject jso = new JuraScriptObject();
            jso.Execute(@"
                var fso = new ActiveXObject(""Scripting.FileSystemObject"");
                var coll = new Enumerator(fso.Drives);
                var t = fso.Drives.Count;
                var drives = [];
                var atEnd = coll.atEnd();
                while(!coll.atEnd()) {
                    drives.push(coll.item());
                    coll.moveNext();
                }
                drives.push(coll.item()); //Add the last item
            ");

            DriveInfo[] di = DriveInfo.GetDrives();

            ArrayInstance driveArr = (ArrayInstance)jso.JurassicEngine.GetGlobalValue("drives");
            
            //uint because GetPropertyValue() requires it
            for (uint i = 0; i < di.Length; i++)
            {
                Assert.AreEqual(driveArr.GetPropertyValue(i).ToString() + "\\", di[i].ToString());
            }
        }

        [TestMethod]
        public void COMTest_FileSystemObject_ReadWriteFiles()
        {
            if (File.Exists("test.txt")) { File.Delete("test.txt"); }
            //Enumerate drives
            JuraScriptObject jso = new JuraScriptObject();
            jso.Execute(@"
                var fso = new ActiveXObject(""Scripting.FileSystemObject"");
                var objFile = fso.CreateTextFile(""test.txt"");
                objFile.Close();
            ");
            Assert.IsTrue(File.Exists("test.txt"));

            jso.Execute(@"
                var fso2 = new ActiveXObject(""Scripting.FileSystemObject"");
                var objFile2 = fso.OpenTextFile(""test.txt"",2); //Open for writing
                objFile2.WriteLine(""sometext"");
                objFile2.Close();
            ");

            Assert.AreEqual("sometext\r\n", File.ReadAllText("test.txt"));

            jso.Execute(@"
                var fso3 = new ActiveXObject(""Scripting.FileSystemObject"");
                var objFile3 = fso.OpenTextFile(""test.txt"",1); //Open for reading
                var textRead = objFile3.ReadLine();
                objFile3.Close();
            ");
            //We only read a line, so the linebreak wont be returned
            Assert.AreEqual(jso.JurassicEngine.GetGlobalValue("textRead").ToString(), "sometext");

        }

        [TestMethod]
        public void COMTest_Excel_ReadWriteCell()
        {
            JuraScriptObject jso = new JuraScriptObject();
            jso.Execute(@"
                var objExcel = new ActiveXObject(""Excel.Application"");
                objExcel.Visible = true;
                objExcel.Workbooks.Add();
                objExcel.Cells(1,1).Value = ""Test Value"";
                objExcel.Cells(""A1"").Font.Bold = true;
                objExcel.Cells(1,1).Font.Size = 24;
                objExcel.Cells(1,1).Font.ColorIndex = 3;
                
            ");
        }


        /*
        /// <summary>
        ///A test for GetCOMMemberType
        ///</summary>
        [TestMethod()]
        public void GetCOMMemberTypeTest()
        {
            //string memberName = string.Empty; // TODO: Initialize to an appropriate value
            //ITypeInfo typeInfo = null; // TODO: Initialize to an appropriate value
            //TYPEATTR typeAttr = new TYPEATTR(); // TODO: Initialize to an appropriate value
            //ComReflector.ComReflectedMemberTypes expected = new ComReflector.ComReflectedMemberTypes(); // TODO: Initialize to an appropriate value
            //ComReflector.ComReflectedMemberTypes actual;
            //actual = ComReflector.GetCOMMemberType(memberName, typeInfo, typeAttr);
            //Assert.AreEqual(expected, actual);
            Assert.Inconclusive("Verify the correctness of this test method.");
        }

        /// <summary>
        ///A test for GetDispId
        ///</summary>
        [TestMethod()]
        public void GetDispIdTest()
        {
            //IDispatch idisp = null; // TODO: Initialize to an appropriate value
            //string memberName = string.Empty; // TODO: Initialize to an appropriate value
            //int expected = 0; // TODO: Initialize to an appropriate value
            //int actual;
            //actual = ComReflector.GetDispId(idisp, memberName);
            //Assert.AreEqual(expected, actual);
            Assert.Inconclusive("Verify the correctness of this test method.");
        }

        /// <summary>
        ///A test for GetMemberType
        ///</summary>
        [TestMethod()]
        public void GetMemberTypeTest()
        {
            //object comObject = null; // TODO: Initialize to an appropriate value
            //string memberName = string.Empty; // TODO: Initialize to an appropriate value
            //ComReflector.ComReflectedMemberTypes expected = new ComReflector.ComReflectedMemberTypes(); // TODO: Initialize to an appropriate value
            //ComReflector.ComReflectedMemberTypes actual;
            //actual = ComReflector.GetMemberType(comObject, memberName);
            //Assert.AreEqual(expected, actual);
            Assert.Inconclusive("Verify the correctness of this test method.");
        }

        /// <summary>
        ///A test for DumpTypeDesc
        ///</summary>
        [TestMethod()]
        public void DumpTypeDescTest()
        {
            //TYPEDESC tdesc = new TYPEDESC(); // TODO: Initialize to an appropriate value
            //ITypeInfo context = null; // TODO: Initialize to an appropriate value
            //string expected = string.Empty; // TODO: Initialize to an appropriate value
            //string actual;
            //actual = ComReflector.DumpTypeDesc(tdesc, context);
            //Assert.AreEqual(expected, actual);
            Assert.Inconclusive("Verify the correctness of this test method.");
        }

        /// <summary>
        ///A test for DefaultPropertyName
        ///</summary>
        [TestMethod()]
        public void DefaultPropertyNameTest()
        {
            //IDispatch idisp = null; // TODO: Initialize to an appropriate value
            //ComReflector.ComReflectedMemberTypes memberType = new ComReflector.ComReflectedMemberTypes(); // TODO: Initialize to an appropriate value
            //ComReflector.ComReflectedMemberTypes memberTypeExpected = new ComReflector.ComReflectedMemberTypes(); // TODO: Initialize to an appropriate value
            //string expected = string.Empty; // TODO: Initialize to an appropriate value
            //string actual;
            //actual = ComReflector.DefaultPropertyName(idisp, out memberType);
            //Assert.AreEqual(memberTypeExpected, memberType);
            //Assert.AreEqual(expected, actual);
            Assert.Inconclusive("Verify the correctness of this test method.");
        }
         * 
         */
    }
}
