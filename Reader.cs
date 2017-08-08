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
        private _Application _excel;
        private Workbooks _wbs;
        private _Workbook _wb;
        private _Worksheet _ws;
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
        public Reader()
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
    }
}
