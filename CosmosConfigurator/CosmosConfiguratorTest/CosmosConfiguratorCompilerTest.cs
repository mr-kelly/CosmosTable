using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using CosmosConfigurator;
using AppConfigs;

namespace CosmosConfiguratorTest
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
                    ExportTabExt = ".bytes",
                    ExportCodePath = "../../TabConfigs.cs",
                });
            Assert.IsTrue(compiler.Compile("./test_excel.xls"));
        }

        [TestMethod]
        public void ReadCompliedTsv()
        {
            var tabFile = TabFile.LoadFromFile("./test_excel.bytes");
            Assert.AreEqual<int>(3, tabFile.GetColumnCount());

            var headerNames = tabFile.HeaderNames.ToArray();
            Assert.AreEqual("Id", headerNames[0]);
            Assert.AreEqual("Name", headerNames[1]);
            Assert.AreEqual("StrArray", headerNames[2]);
        }

        [TestMethod]
        public void ReadCompliedTsvWithClass()
        {
            var tabFile = TabFile<TestExcelConfig>.LoadFromFile("./test_excel.bytes");

            var config = tabFile.FindByPrimaryKey(1);

            Assert.IsNotNull(config);
            Assert.AreEqual(config.Name, "Test1");
        }


        /// <summary>
        /// 测试写入TSV
        /// </summary>
        [TestMethod]
        public void TestWriteTSV()
        {

        }
    }
}
