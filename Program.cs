using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace ConsoleApplication1
{
    /// <summary>
    /// 
    /// Project entry point. 
    /// 
    /// @author: Alexander James Bochel
    /// @version: 7/5/2017
    /// 
    /// </summary>
    class Program
    {
        static Reader reader;
        static Retreiver retreiver;

        static void Main(string[] args)
        {
            openAndExecute();
            print();
        }

        private static void openAndExecute()
        {
            retreiver = new Retreiver();
            retreiver.runOREP();
            
            reader = new Reader();
            reader.modifier.execute();
        }

        private static void print()
        {
            Printer printer = new Printer(reader.modifier.salesList,
                reader.excel, reader.wbs, reader.wb, reader.ws);
        }
    }
}