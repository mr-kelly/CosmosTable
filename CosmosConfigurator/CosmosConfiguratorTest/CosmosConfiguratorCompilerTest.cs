using System;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using CosmosConfigurator;

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
                    ExcelPath = "./test_excel.xls",
                    Ext = ".bytes",
                });
            Assert.IsTrue(compiler.Run());
        }

        [TestMethod]
        public void ReadCompliedTsv()
        {
            var reader = Reader.LoadFromFile("./test_excel.bytes");
            Assert.AreEqual<int>(2, reader.GetColumnCount());

            var headerNames = reader.HeaderNames.ToArray();
            Assert.AreEqual("Id", headerNames[0]);
            Assert.AreEqual("Name", headerNames[1]);
        }
    }
}
