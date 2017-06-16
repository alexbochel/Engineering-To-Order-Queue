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

        

        // Declare list of Sales to be tested. 
        List<Sales> testList = new List<Sales>();

        
        [TestMethod]
        public void TestSort()
        {            
            // Add material names to the sales fields created. 
            sale1.material = "DEF123";
            sale2.material = "ABC123";
            sale3.material = "ABC456";
            sale4.material = "GHI123";
            sale5.material = "GHI223";
            sale6.material = "ZZZ999";
            sale7.material = "ABC124";
            sale8.material = "JKL111";
            sale9.material = "ACB123";
            sale10.material = "ABC124";

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

            // Declare Modifier Object.
            Modifier modifier = new Modifier(testList);

            // Sort sales 1-10 by alphabetical order. 
            modifier.sortMaterial();
            testList = modifier.salesList;

            // Print statements to view whole list in output screen
            Console.WriteLine(testList[0].material.ToString());
            Console.WriteLine(testList[1].material.ToString());
            Console.WriteLine(testList[2].material.ToString());
            Console.WriteLine(testList[3].material.ToString());
            Console.WriteLine(testList[4].material.ToString());
            Console.WriteLine(testList[5].material.ToString());
            Console.WriteLine(testList[6].material.ToString());
            Console.WriteLine(testList[7].material.ToString());
            Console.WriteLine(testList[8].material.ToString());
            Console.WriteLine(testList[9].material.ToString());


            Assert.AreEqual("ABC123", testList[0].material);
            Assert.AreEqual("ABC124", testList[1].material);
            Assert.AreEqual("ABC124", testList[2].material);
            Assert.AreEqual("ZZZ999", testList[9].material);  
        }

        [TestMethod]
        public void TestDeleteFifteens()
        {
            sale1.MSPS = "10";
            sale2.MSPS = "15";
            sale3.MSPS = "15";
            sale4.MSPS = "14";
            sale5.MSPS = "15";

            testList.Add(sale1);
            testList.Add(sale2);
            testList.Add(sale3);
            testList.Add(sale4);
            testList.Add(sale5);

            Modifier modifier = new Modifier(testList);

            // Ensure everything was added correctly.
            Assert.IsTrue(testList.Count == 5);

            modifier.deleteFifteens();

            Assert.IsTrue(testList.Count == 2); 
        }

        [TestMethod]
        public void TestDeleteTJ()
        {
            sale1.material = "ABJFSJTJ"; // Contains TJ
            sale2.material = "KTJANDKFH"; // Contains TJ
            sale3.material = "SADVUAWDJ"; // No TJ
            sale4.material = "JTNDBAYK"; // No TJ

            testList.Add(sale1);
            testList.Add(sale2);
            testList.Add(sale3);
            testList.Add(sale4);

            Modifier modifier = new Modifier(testList);

            Assert.IsTrue(testList.Count == 4);

            modifier.deleteTJ();

            Assert.IsTrue(testList.Count == 0);
        }

        [TestMethod]
        public void testDaysInQueue()
        {

        }
    }
}
