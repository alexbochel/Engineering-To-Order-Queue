using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Reflection;
using System.IO;
using System.Diagnostics;

namespace ConsoleApplication1
{
    /// <summary>
    /// This class retrieves the material and BOM for each material!
    /// </summary>
    public class Retriever
    {
        public List<Sales> salesCompare;

        public _Application excel;
        public Workbooks wbs;
        public _Workbook wb;
        public _Worksheet ws;
        public int sheetNumber = 2;

        /// <summary>
        /// TODO
        /// </summary>
        /// <param name="salesList"></param>
        /// <param name="excel"></param>
        /// <param name="wbs"></param>
        /// <param name="wb"></param>
        /// <param name="ws"></param>
        public Retriever(List<Sales> salesList, _Application excel, Workbooks wbs, _Workbook wb, _Worksheet ws)
        {
            this.excel = excel;
            this.wbs = wbs;
            this.wb = wb;
            this.ws = this.wb.Worksheets.Add(After: this.wb.Sheets[this.wb.Sheets.Count]);
            
            salesCompare = salesList;

            // Runs VBS to copy the BOM data and then adds the BOM
            runZSE16N();
            addBOM();
        }

        /// <summary>
        /// This method calls the embedded VBScript file and runs it before executing the 
        /// rest of the program. 
        /// </summary>
        private static void runZSE16N()
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

        private void addBOM()
        {
            int column = 1;
            int range = 1;

            ws.Paste();

            for (int row = 1; ws.Cells[row, column].Value2 != null; row++)
            {
                string compare = readCell(row, column);
                range++;

                for (int i = 0; i < salesCompare.Count; i++)
                {
                    if (salesCompare[i].material == compare)
                    {
                        salesCompare[i].BOM = "yes";
                    }
                }
            }

            var ranger = ws.get_Range("A1", "D" + range.ToString());
            ranger.Select();
            ranger.Clear();
        }

        private string readCell(int i, int j)
        {
            if (ws.Cells[i, j].Value2 != null)
            {
                string cell = ws.Cells[i, j].Value2.ToString();

                return cell;
            }
            else
            {
                return "";
            }
        }
    }
}
