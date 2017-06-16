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
    /// This class prints the data in the List of Sales onto a new Excel Workbook.
    /// 
    /// Author: Alexander James Bochel
    /// Date Updated: 6/16/2017
    /// 
    /// </summary>
    public class Printer
    {
        public List<Sales> salesList { get; set; }
        public string path { get; set; }
        public _Application excel;
        public Workbooks wbs;
        public _Workbook wb;
        public _Worksheet ws;
        public int sheetNumber = 2;

        public Printer(List<Sales> salesList, _Application excel, Workbooks wbs, _Workbook wb, _Worksheet ws)
        {
            this.salesList = salesList;

            this.excel = excel;
            this.wbs = wbs;
            this.wb = wb;
            this.ws = wb.Worksheets[sheetNumber];

            printColumnNames();
            printRows();
            ws.Columns.AutoFit();
            this.ws.Columns.AutoFit();
            garbageCleanup();
        }

        /// <summary>
        /// This method prints the Sales fields in each column and leaves column D 
        /// blank. 
        /// </summary>
        public void printRows()
        {
            int lastRow = 1;

            for (int row = 2; row <= salesList.Count + lastRow; row++)
            {
                int cellHoriz = 1;

                printCell(row, cellHoriz, salesList[row - 2].salesNum);
                cellHoriz++;

                printCell(row, cellHoriz, salesList[row - 2].material);
                cellHoriz++;

                printCell(row, cellHoriz, salesList[row - 2].description);
                cellHoriz++;
                cellHoriz++; // In order to keep a blank column. 

                printCell(row, cellHoriz, salesList[row - 2].MSPS);
                cellHoriz++;

                printCell(row, cellHoriz, salesList[row - 2].MRPC);
                cellHoriz++;

                printCell(row, cellHoriz, salesList[row - 2].quantity);
                cellHoriz++;

                printCell(row, cellHoriz, salesList[row - 2].formatDate.ToString("MM/dd/yyyy")); // Prints date
                cellHoriz++;

                printCell(row, cellHoriz, salesList[row - 2].daysInQueue.ToString()); // Prints days in queue
            }
        }

        /// <summary>
        /// This method prints the names of the columns on the first row. 
        /// </summary>
        public void printColumnNames()
        {
            printCell(1, 1, "Sales");
            printCell(1, 2, "Material");
            printCell(1, 3, "Description");
            printCell(1, 4, "BOMs");
            printCell(1, 5, "MS-PS");
            printCell(1, 6, "MRPC");
            printCell(1, 7, "Quantity");
            printCell(1, 8, "Date");
            printCell(1, 9, "Days in Queue");
        }

        /// <summary>
        /// This method realeses and closes the COM's in order to allow excel to close. 
        /// </summary>
        public void garbageCleanup()
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
