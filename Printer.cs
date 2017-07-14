using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;


namespace ConsoleApplication1
{
    /// <summary>
    /// 
    /// This class prints the data in the List of Sales onto a new Excel Workbook.
    /// 
    /// @author: Alexander James Bochel
    /// @version: 7/14/2017
    /// 
    /// </summary>
    public class Printer
    {
        public List<Sales> salesList { get; set; }
        public string path { get; set; }
        public Retriever retriever;
        public _Application excel;
        public Workbooks wbs;
        public _Workbook wb;
        public _Worksheet ws;
        public int sheetNumber = 2;

        /// <summary>
        /// This constructor takes all of the excel information from the reader class that creates it. 
        /// </summary>
        /// <param name="salesList"> List of sales read by reader. </param>
        /// <param name="excel"> Application instance created by reader. </param>
        /// <param name="wbs"> Workbooks created by reader. </param>
        /// <param name="wb"> Workbook created by reader. </param>
        /// <param name="ws"> Worksheet created by reader. </param>
        public Printer(List<Sales> salesList, _Application excel, Workbooks wbs, _Workbook wb, _Worksheet ws)
        {
            this.salesList = salesList;

            this.excel = excel;
            this.wbs = wbs;
            this.wb = wb;
            this.ws = wb.Worksheets[sheetNumber];

            // This section handles getting the material to be pasted into SAP. 
            tempPrintMat();
            salesList = retriever.salesCompare;

            // Printing and cleanup.
            print();
            ws.Columns.AutoFit();
            garbageCleanup();
        }

        /// <summary>
        /// This handles the final printing data and formatting of the cells. 
        /// </summary>
        public void print()
        {
            printColumnNames();
            int lastRow = printRows();
            this.ws.Columns.AutoFit();
            addBorders(lastRow);
            this.wb.Worksheets[2].Select();
        }

        private void addBorders(int numRows)
        {
            var range = ws.get_Range("A18", "M" + numRows);
            range.Borders.LineStyle = XlLineStyle.xlContinuous;
        }

        /// <summary>
        /// This method prints the Sales fields in each column and leaves column D 
        /// blank.
        /// </summary>
        /// <returns> Final row number that was printed. </returns>
        public int printRows()
        {
            int lastRow = 18;
            int numRow = 19;

            for (int row = 19; row <= salesList.Count + lastRow; row++)
            {
                int cellHoriz = 1;

                printCell(row, cellHoriz, salesList[row - 19].salesNum);
                cellHoriz++;

                printCell(row, cellHoriz, salesList[row - 19].material);
                cellHoriz++;

                printCell(row, cellHoriz, salesList[row - 19].description);
                cellHoriz++;

                printCell(row, cellHoriz, salesList[row - 19].BOM);
                cellHoriz++;  

                printCell(row, cellHoriz, salesList[row - 19].MSPS);
                cellHoriz++;

                printCell(row, cellHoriz, salesList[row - 19].MRPC);
                cellHoriz++;

                printCell(row, cellHoriz, salesList[row - 19].quantity);
                cellHoriz++;

                printCell(row, cellHoriz, salesList[row - 19].formatDate.ToString("MM/dd/yyyy")); // Prints date
                cellHoriz = cellHoriz + 5;

                printCell(row, cellHoriz, salesList[row - 19].daysInQueue.ToString()); // Prints days in queue

                numRow = row;
            }

            return numRow;
        }

        /// <summary>
        /// This method temporarily prints the materials so that they can be copied and pasted into SAP.
        /// </summary>
        private void tempPrintMat()
        {
            int column = 1;
            int finalRow = 1;

            for (int row = 1; row <= salesList.Count; row++)
            {
                printCell(row, column, salesList[row - 1].material);
                finalRow++;
            }

            ws.Activate();

            var range = ws.get_Range("A1", "A" + finalRow.ToString());
            range.Select();
            range.Copy();
            retriever = new Retriever(salesList, excel, wbs, wb, ws);
            range.Clear();
        }

        /// <summary>
        /// This method prints the names of the columns on the first row. 
        /// </summary>
        private void printColumnNames()
        {
            printCell(18, 1, "Sales");
            printCell(18, 2, "Material");
            printCell(18, 3, "Description");
            printCell(18, 4, "BOMs");
            printCell(18, 5, "MS-PS");
            printCell(18, 6, "MRPC");
            printCell(18, 7, "Quantity");
            printCell(18, 8, "Created On");
            printCell(18, 13, "Days in Queue");
        }

        /// <summary>
        /// This method realeses and closes the COM's in order to allow excel to close. 
        /// </summary>
        private void garbageCleanup()
        {
            wb.Close();
            wbs.Close();
            excel.Quit();

            Marshal.ReleaseComObject(wb);
            Marshal.ReleaseComObject(wbs);
            Marshal.ReleaseComObject(excel);
        }

        private void printCell(int i, int j, string value)
        {
            ws.Cells[i, j].Value2 = value;             
        }
    }
}
