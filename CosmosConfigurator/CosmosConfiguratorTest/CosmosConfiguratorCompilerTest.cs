using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using CosmosConfigurator;

namespace CosmosConfiguratorTest
{
    [TestClass]
    public class CosmosConfiguratorCompilerTest
    {
        [TestMethod]
        public void TestCompileExcelFile1()
        {
            var compiler = new Compiler("./test_excel.xls");
            Assert.IsTrue(compiler.Run());
        }
    }
}
