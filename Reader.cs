using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Windows;

namespace ConsoleApplication1
{
    /// <summary>
    /// 
    /// This class reads data from an excel sheet and stores the data in Sales objects
    /// and then places each Sale into a List of Sales. 
    /// 
    /// @author: Alexander James Bochel
    /// @version: 7/14/2017
    /// 
    /// </summary>
    public class Reader
    {
        private int sheet = 1;
        private int carryOver = 0;
        private int buffer = 0;
        private _Application _excel;
        private Workbooks _wbs;
        private _Workbook _wb;
        private _Worksheet _ws;
        private _Application _prevExcel;
        private Workbooks _prevWbs;
        private _Workbook _prevWb;
        private _Worksheet _prevWs;
        private Modifier _modifier;
        private List<Sales> _salesList;
        
        /// <summary>
        /// Getter/Setter: The modifier object that manipulates the data from SAP. 
        /// </summary>
        public Modifier modifier 
        {
            get { return _modifier; }
            set { _modifier = value; }
        }

        /// <summary>
        /// Getter/Setter: The excel instance with the report in it. 
        /// </summary>
        public _Application excel
        {
            get { return _excel; }
            set { _excel = value; }
        }

        /// <summary>
        /// Getter/Setter: Excel Workbooks. 
        /// </summary>
        public Workbooks wbs
        {
            get { return _wbs; }
            set { _wbs = value; }
        }

        /// <summary>
        /// Getter/Setter: The workbook open in excel. 
        /// </summary>
        public _Workbook wb
        {
            get { return _wb; }
            set { _wb = value; }
        }

        /// <summary>
        /// Getter/Setter: The worksheet number. 
        /// </summary>
        public _Worksheet ws
        {
            get { return _ws; }
            set { _ws = value; }
        }

        /// <summary>
        /// Getter/Setter: The excel instance with the report in it. 
        /// </summary>
        public _Application prevExcel
        {
            get { return _prevExcel; }
            set { _prevExcel = value; }
        }

        /// <summary>
        /// Getter/Setter: Excel Workbooks. 
        /// </summary>
        public Workbooks prevWbs
        {
            get { return _prevWbs; }
            set { _prevWbs = value; }
        }

        /// <summary>
        /// Getter/Setter: The workbook open in excel. 
        /// </summary>
        public _Workbook prevWb
        {
            get { return _prevWb; }
            set { _prevWb = value; }
        }

        /// <summary>
        /// Getter/Setter: The worksheet number. 
        /// </summary>
        public _Worksheet prevWs
        {
            get { return _prevWs; }
            set { _prevWs = value; }
        }

        /// <summary>
        /// Getter/Setter: The list of sales created based off of the rows of SAP data. 
        /// </summary>
        public List<Sales> salesList
        {
            get { return _salesList; }
            set { _salesList = value; }
        }

        /// <summary>
        /// This constructor opens a new excel file and copies the clipboard onto 
        /// the excel file. It also creates a list of sales and creates a modifier 
        /// object. 
        /// </summary>
        public Reader(string prevPath)
        {
            Console.WriteLine("Reading SAP data...");
            
            excel = new _Excel.Application();
            wbs = excel.Workbooks;             // Easier garbage cleanup when split up. 
            wb = excel.Workbooks.Add();
            
            // Ensure that there are enough worksheets. 
            while (wb.Worksheets.Count < 3)
            {
                wb.Worksheets.Add(After: wb.Sheets[wb.Sheets.Count]);
            }
            
            ws = wb.Worksheets[sheet];
            ws.Paste();

            salesList = new List<Sales>();
            createSales();
            addAdditional(prevPath);

            modifier = new Modifier(salesList);
        }   

        /// <summary>
        /// This method creates a new Sale for every row in the excel file. 
        /// </summary>
        /// <returns> Number of rows (sales) in the excel sheet. </returns>
        public int createSales()
        {
            int rows = 1; // Excel sheets start at 1 not 0. 

            while (ws.Cells[rows, 1].Value2 != null)
            {
                Sales sale = new Sales();
                addFields(sale, rows);
                salesList.Add(sale);

                rows++;
            }

            return rows;
        }

        /// <summary>
        /// This helper method adds fields to all of the sales.
        /// </summary>
        /// <param name="sale"> Sale that is getting fields filled. </param>
        /// <param name="row"> Row to look for fields on </param>
        public void addFields(Sales sale, int row)
        {
            int i = 1;

            sale.salesNum = readCell(row, i);              // Sales Number field
            i++;

            sale.material = readCell(row, i);              // material field
            i++;

            sale.description = readCell(row, i);           // Description field
            i++;

            sale.MSPS = readCell(row, i);                  // MSPS field
            i++;

            sale.MRPC = readCell(row, i);                  // MRPC field
            i++;

            sale.quantity = readCell(row, i);              // Quantity field
            i++;

            sale.date = readCell(row, i);                  // Date field            
        }

        /// <summary>
        /// Adds information from previous days ETO queue. 
        /// </summary>
        private void addAdditional(string prevPath)
        {
            int j = 1;
            int i;
            int prevMatCol = 2;
            int prevDescCol = 3;
            int prevMRPCCol = 6;

            openPrevReport(prevPath);

            for (int k = 0; k < salesList.Count; k++)
            {
                i = 19;

                while (readPrevCell(i, j) != "")
                {
                    string previousMat;
                    string previousDesc;
                    string previousMRPC;

                    previousMRPC = readPrevCell(i, prevMRPCCol);
                    previousMat = readPrevCell(i, prevMatCol);
                    previousDesc = readPrevCell(i, prevDescCol);
                    
                    if (previousMat == salesList[k].material && previousDesc == salesList[k].description)
                    {
                        copyData(k, i);
                        checkNewMRPC(previousMRPC, i, k);
                    }

                    i++;
                }
                
                
                // Check if there is a carry over between reports. 
                //if (checkCarryOver(readPrevCell(i, prevMatCol),
                //    readPrevCell(i, prevDescCol)))
                //{
                //    checkNewMRPC(readPrevCell(i, prevMRPCCol), i);
                //    copyData(i);
                //}
                //i++;
            }
        }

        /// <summary>
        /// Copies the notes and rankings from the old excel sheet into the sales list. 
        /// </summary>
        /// <param name="i"> Current row number. </param>
        private void copyData(int k, int i)
        {
            int addNotesCol = 10;
            int engeNotesCol = 11;
            int rankingCol = 12;

            salesList[k].addNotes = readPrevCell(i, addNotesCol);
            salesList[k].engeNotes = readPrevCell(i, engeNotesCol);
            salesList[k].ranking = readPrevCell(i, rankingCol);
        }

        /// <summary>
        /// Checks to see if the MRPC has changed since last time. 
        /// </summary>
        /// <param name="prevMRPC"></param>
        /// <param name="row"></param>
        private void checkNewMRPC(string prevMRPC, int row, int k)
        {
            int cBurgDateCol = 9;

            // Check for the same MRPC
            if(prevMRPC == salesList[k].MRPC)
            {
                salesList[k].cBurgDate = readPrevCell(row, cBurgDateCol);
            }
            else if (prevMRPC == "BP2" && salesList[k].MRPC == "BP3")
            {
                salesList[k].cBurgDate = DateTime.Today.ToString("MM/dd/yyyy");
            }
            else if (prevMRPC == "13" && salesList[k].MRPC == "ETO")
            {
                salesList[k].cBurgDate = DateTime.Today.ToString("MM/dd/yyyy");
            }
        }

        //private bool checkCarryOver(string previousMat, string previousDesc)
        //{
        //    for (int i = 0 + buffer; i < salesList.Count; i++)
        //    {
        //        if (previousMat == salesList[i].material && previousDesc == salesList[i].description) // AHHHH: This misses subsequent matNumbms with the same value. 
        //        {
        //            carryOver = i;
        //            buffer++;
        //            return true;
        //        }
        //    }

        //    return false;
        //}

        /// <summary>
        /// Opens a dialog box that allows the user to choose the last ETO report. 
        /// </summary>
        private void openPrevReport(string prevPath) // TODO: Give this the check for correct worksheet. 
        {
            int sheetNum = 2;

            // New instance of excel to read previous sheet's data. 
            prevWbs = excel.Workbooks;
            prevWb = prevWbs.Open(prevPath);
            prevWs = prevWb.Worksheets[sheetNum];
        }

        /// <summary>
        /// This method reads in a cell from excel. 
        /// </summary>
        /// <param name="i"> The "y" coordinate. </param>
        /// <param name="j"> The "x" coordinate. </param>
        /// <returns> The value in the cell as a string. </returns>
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

        /// <summary>
        /// This method reads in a cell from excel. 
        /// </summary>
        /// <param name="i"> The "y" coordinate. </param>
        /// <param name="j"> The "x" coordinate. </param>
        /// <returns> The value in the cell as a string. </returns>
        private string readPrevCell(int i, int j)
        {
            if (prevWs.Cells[i, j].Value2 != null)
            {
                string cell = prevWs.Cells[i, j].Value2.ToString();

                return cell;
            }
            else
            {
                return "";
            }
        }
    }
}
