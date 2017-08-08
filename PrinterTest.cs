using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections;
using System.Collections.Generic;
using ConsoleApplication1;

namespace AutomationProgramTests
{
    [TestClass]
    public class PrinterTest
    {
        Reader reader;
        Printer printer;

        /// <summary>
        /// This test method simply opens excel so that the printing process is visible. 
        /// </summary>
        [TestMethod]
        public void observePrint()
        { 
            try
            {
                reader = new Reader();
                reader.modifier.execute();
                printer = new Printer(reader.salesList, reader.excel, reader.wbs, reader.wb, reader.ws);
            }
            catch (Exception e)
            {
                Assert.IsTrue(e != null);
                Assert.IsTrue(e is FormatException);
            }
        }
    }
}
