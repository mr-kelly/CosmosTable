using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using AppConfigs;

namespace CosmosTable.Test
{
    [TestClass]
    public class CosmosConfiguratorCompilerTest
    {
        [TestMethod]
        public void CompileTestExcel()
        {
            var compiler = new Compiler(
                new CompilerConfig
                {
                    CodeTemplates = new Dictionary<string, string>()
                    {
                        {File.ReadAllText("./GenCode.cs.tpl"), "../../TabConfigs.cs"},  // code template -> CodePath
                    },
                    ExportTabExt = ".bytes",
                    
                });
            Assert.IsTrue(compiler.Compile("./test_excel.xls"));
        }

        [TestMethod]
        public void ReadCompliedTsv()
        {
            var tabFile = TableFile.LoadFromFile("./test_excel.bytes");
            Assert.AreEqual<int>(3, tabFile.GetColumnCount());

            var headerNames = tabFile.HeaderNames.ToArray();
            Assert.AreEqual("Id", headerNames[0]);
            Assert.AreEqual("Name", headerNames[1]);
            Assert.AreEqual("StrArray", headerNames[2]);
        }

        [TestMethod]
        public void ReadCompliedTsvWithClass()
        {
            var tabFile = TableFile<TestExcelInfo>.LoadFromFile("./test_excel.bytes");

            var config = tabFile.FindByPrimaryKey(1);

            Assert.IsNotNull(config);
            Assert.AreEqual(config.Name, "Test1");
        }



        class TestWrite : TableRowInfo
        {
            public override bool IsAutoParse
            {
                get { return true; }
            }

            public string TestColumn1;
            public int TestColumn2;
        }

        /// <summary>
        /// 测试写入TSV
        /// </summary>
        [TestMethod]
        public void TestWriteTableFile1()
        {
            var tabFileWrite = new TabFileWriter<TestWrite>();
            var newRow = tabFileWrite.NewRow();
            newRow.TestColumn1 = "Test String";
            newRow.TestColumn2 = 123123;

            tabFileWrite.Save("./test_write.bytes");


            var tabFileRead = TableFile<TestWrite>.LoadFromFile("./test_write.bytes");
            Assert.AreEqual(tabFileRead.GetHeight(), 1);

        }

        /// <summary>
        /// 读入，然后再写入测试
        /// </summary>
        [TestMethod]
        public void TestWriteTableFile2()
        {
            var tabFile = TableFile<TestWrite>.LoadFromFile("./test_write.bytes");

            var tabFileWrite = new TabFileWriter<TestWrite>(tabFile);

            var newRow = tabFileWrite.NewRow();
            newRow.TestColumn1 = Path.GetRandomFileName();
            newRow.TestColumn2 = new Random().Next();


            // 两个方法执行后
            Assert.AreEqual(tabFile.GetHeight(), 2);

            tabFileWrite.Save("./test_write.bytes");

        }
    }
}
