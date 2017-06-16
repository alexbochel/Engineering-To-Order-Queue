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

        [TestMethod]
        public void observePrint()
        {
            reader = new Reader();
            reader.modifier.execute();
            printer = new Printer(reader.salesList, reader.excel, reader.wbs, reader.wb, reader.ws);
        }
    }
}
