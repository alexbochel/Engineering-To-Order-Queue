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
    public class Retreiver
    {

        private List<Sales> _salesCompare;
        private _Application excel;
        private Workbooks wbs;
        private _Workbook wb;
        private _Worksheet ws;
        private int _numHits;

        /// <summary>
        /// Getter/Setter: Sales to be entered into the BOM
        /// </summary>
        public List<Sales> salesCompare
        {
            get { return _salesCompare; }
            set { _salesCompare = value; }
        }

        /// <summary>
        /// Getter/Setter: Number of materials with a BOM. 
        /// </summary>
        public int numHits
        {
            get { return _numHits; }
            set { _numHits = value; }
        }

        /// <summary>
        /// This constructor adds a new sheet to the excel workbook, runs a section of VBScript, 
        /// </summary>
        /// <param name="salesList"> List of sales from SAP. </param>
        /// <param name="excel"> Instance of excel. </param>
        /// <param name="wbs"> Workbooks from the printer/reader. </param>
        /// <param name="wb"> Workbook from the printer/reader. </param>
        /// <param name="ws"> Worksheet being printed on. </param>
        public Retreiver(List<Sales> salesList, _Application excel, Workbooks wbs, _Workbook wb, _Worksheet ws)
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
        /// Default constructor for the retreiver class. 
        /// </summary>
        public Retreiver()
        {
            Console.WriteLine("Waiting for SAP Queue Data...");
        }
        
        /// <summary>
        /// This method calls the embedded VBScript file and runs OREP before executing the 
        /// rest of the program. 
        /// </summary>
        public void runOREP()
        {
            var assembly = Assembly.GetExecutingAssembly();
            //Getting names of all embedded resources
            var allResourceNames = assembly.GetManifestResourceNames();
            //Selecting first one. 
            var resourceName = allResourceNames[1];
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

        /// <summary>
        /// This method calls the embedded VBScript file and runs ZSE16N before executing the 
        /// rest of the program. 
        /// </summary>
        private static void runZSE16N()
        {
            Console.WriteLine("Waiting for BOM Data...");
            
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

        /// <summary>
        /// Adds BOM data to each sale in the sales list and then clears the pasted data from the third worksheet. 
        /// </summary>
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

            numHits = range - 1;

            string cellA = "A1:D" + range.ToString();
            ws.get_Range(cellA, Type.Missing).Select();
            ws.get_Range(cellA, Type.Missing).Clear();
            //var ranger = ws.get_Range("A1", "D" + range.ToString());
            //ranger.Select();
            //ranger.Clear();
        }

        /// <summary>
        /// Reads in a cell from excel. 
        /// </summary>
        /// <param name="i"> The "y" coordinate. </param>
        /// <param name="j"> The "x" coordinate. </param>
        /// <returns> The value of the cell as a string. </returns>
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
