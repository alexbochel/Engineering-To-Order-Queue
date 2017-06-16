using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace ConsoleApplication1
{
    /// <summary>
    /// Author: Alexander James Bochel
    /// Date Updated: 6/16/2017
    /// </summary>
    class Program
    {
        static Reader reader;

        /// <summary>
        /// This will call the rest of the classes in the program. 
        /// </summary>
        /// <param name="args"> Command line arguments. </param>
        static void Main(string[] args)
        {
            openAndExecute();
            print();
        }

        /// <summary>
        /// This method creates a reader that automatically reads in data upon creation and also
        /// has a function call to the modifier to begin modifying the list of Sales. 
        /// </summary>
        public static void openAndExecute()
        {
            reader = new Reader();
            reader.modifier.execute();
        }

        /// <summary>
        /// This method creates the printer object that automatically starts printing upon creation. 
        /// </summary>
        public static void print()
        {
            Printer printer = new Printer(reader.modifier.salesList, 
                reader.excel, reader.wbs, reader.wb, reader.ws);

        }         
    }
}
