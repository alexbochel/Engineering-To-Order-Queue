using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections;
using System.Collections.Generic;
using ConsoleApplication1;

namespace AutomationProgramTests
{
    [TestClass]
    public class ModifierTest
    {
        // Declare list of Sales to be tested. 
        List<Sales> testList = new List<Sales>();
        
        // Declare ten sales to be used in the modifier for the tests.
        Sales sale1 = new Sales();
        Sales sale2 = new Sales();
        Sales sale3 = new Sales();
        Sales sale4 = new Sales();
        Sales sale5 = new Sales();
        Sales sale6 = new Sales();
        Sales sale7 = new Sales();
        Sales sale8 = new Sales();
        Sales sale9 = new Sales();
        Sales sale10 = new Sales();

        /// <summary>
        /// Ensures that every test begins with this setup of data. 
        /// </summary>
        [TestInitialize]
        public void setUp()
        {
            // Set the information for each sale. 
            sale1.material = "KVQWERT";                 // Deleted
            sale1.MSPS = "10";

            sale2.material = "KVQWERTY";
            sale2.MSPS = "15";

            sale3.material = "FGAHJWI";                 // Deleted
            sale3.MSPS = "10";

            sale4.material = "CQWERTY";
            sale4.MSPS = "10";

            sale5.material = "QWERTY";                  // Deleted
            sale5.MSPS = "15";

            sale6.material = "PRESQWERTY";              // Deleted
            sale6.MSPS = "10";

            sale7.material = "AQWERTYUIOP";
            sale7.MSPS = "10";

            sale8.material = "QWERTYJ";                 // Deleted
            sale8.MSPS = "10";

            sale9.material = "ZQWERT";                   // Deleted
            sale9.MSPS = "15";

            sale10.material = "DQWJERTY";                // Deleted
            sale10.MSPS = "15";
            
            // Add sales to the list of sales. 
            testList.Add(sale1);
            testList.Add(sale2);
            testList.Add(sale3);
            testList.Add(sale4);
            testList.Add(sale5);
            testList.Add(sale6);
            testList.Add(sale7);
            testList.Add(sale8);
            testList.Add(sale9);
            testList.Add(sale10);
        }

        /// <summary>
        /// This method tests the sorting method in Modifier.
        /// </summary>
        [TestMethod]
        public void TestSort()
        {
            // Declare Modifier Object.
            Modifier modifier = new Modifier(testList);

            // Sort sales 1-10 by alphabetical order. 
            modifier.sortMaterial();
            testList = modifier.salesList;

            Assert.AreEqual("AQWERTYUIOP", testList[0].material);
            Assert.AreEqual("CQWERTY", testList[1].material);
            Assert.AreEqual("DQWJERTY", testList[2].material);
            Assert.AreEqual("ZQWERT", testList[9].material);
        }

        /// <summary>
        /// This test checks to see if the correct entries are deleted form the list in Modifier class. 
        /// </summary>
        [TestMethod]
        public void TestModifyEntries()
        {
            Modifier modifier = new Modifier(testList);

            modifier.modifyEntries();
            testList = modifier.salesList;

            Assert.AreEqual(3, testList.Count);
            Assert.AreEqual("KVQWERTY", testList[0].material);
            Assert.AreEqual("CQWERTY", testList[1].material);
            Assert.AreEqual("AQWERTYUIOP", testList[2].material);
        }

    }  
}
