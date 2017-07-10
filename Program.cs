using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;

namespace ConsoleApplication1
{
    /// <summary>
    /// Project entry point. 
    /// 
    /// @author: Alexander James Bochel
    /// @version: 7/5/2017
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
            runSAP();
            openAndExecute();
            print();
        }

        /// <summary>
        /// This method creates a reader that automatically reads in data upon creation and also
        /// has a function call to the modifier to begin modifying the list of Sales. 
        /// </summary>
        private static void openAndExecute()
        {
            reader = new Reader();
            reader.modifier.execute();
        }

        /// <summary>
        /// This method creates the printer object that automatically starts printing upon creation. 
        /// </summary>
        private static void print()
        {
            Printer printer = new Printer(reader.modifier.salesList,
                reader.excel, reader.wbs, reader.wb, reader.ws);
        }

        /// <summary>
        /// This method calls the embedded VBScript file and runs it before executing the 
        /// rest of the program. 
        /// </summary>
        private static void runSAP()
        {
            var assembly = Assembly.GetExecutingAssembly();
            //Getting names of all embedded resources
            var allResourceNames = assembly.GetManifestResourceNames();
            //Selecting first one. 
            var resourceName = allResourceNames[0];
            var pathToFile = Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory) +
                              resourceName;

            using (var stream = assembly.GetManifestResourceStream(resourceName))
            using (var fileStream = File.Create(pathToFile))
            {
                stream.Seek(0, SeekOrigin.Begin);
                stream.CopyTo(fileStream);
            }

            Process process = new Process();
            process.StartInfo.FileName = pathToFile;
            process.Start();
            process.WaitForExit();
            process.Close();
        }
    }
}