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
    /// @version: 8/8/2017
    /// 
    /// </summary>
    public class Printer
    {
        private Retreiver retriever;
        private _Application excel;
        private Workbooks wbs;
        private _Workbook wb;
        private _Worksheet ws;
        private const int sheetNumber = 2;
        private int sumDuration = 0;
        private int numOrders = 0;

        /// <summary>
        /// List of sales for determining if individual sales have a BOM. 
        /// </summary>
        public List<Sales> salesList { get; set; }

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
            // Assign the parameters. 
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
            Console.WriteLine("Printing Report...");
            
            printColumnNames();
            int lastRow = printRows();   // Prints the SAP data. 
            printTable();
            printProcess();
            printOverall();
            this.ws.Columns.AutoFit();
            printHiddenColumns();
            addBorders(lastRow);
            addColor();
            addFormulas();
            addGraph();
            formatAlignment();
            
            // Leaves main worksheet open. 
            this.wb.Worksheets[2].Select();
        }

        /// <summary> 
        /// This method adds borders to the excel sheet. 
        /// </summary>
        /// <param name="numRows"> The number of rows that need borders. </param>
        private void addBorders(int numRows)
        {
            var range = ws.get_Range("A18", "AY" + numRows);
            range.Borders.LineStyle = XlLineStyle.xlContinuous;

            range = ws.get_Range("B6", "G17");
            range.Borders.LineStyle = XlLineStyle.xlContinuous;

            range = ws.get_Range("J2", "J3");
            range.Borders.LineStyle = XlLineStyle.xlContinuous;

            range = ws.get_Range("J6", "J7");
            range.Borders.LineStyle = XlLineStyle.xlContinuous;

            range = ws.get_Range("J9", "J10");
            range.Borders.LineStyle = XlLineStyle.xlContinuous;

            range = ws.get_Range("J12", "J14");
            range.Borders.LineStyle = XlLineStyle.xlContinuous;

            range = ws.get_Range("L2", "M3");
            range.Borders.LineStyle = XlLineStyle.xlContinuous;

            range = ws.get_Range("L5", "L6");
            range.Borders.LineStyle = XlLineStyle.xlContinuous;
        }

        /// <summary>
        /// This prints the overview table 
        /// </summary>
        private void printTable()
        {
            // Table Header "wip". 
            printCell(6, 4, "WIP");
            ws.get_Range("D6", "F6").Merge(); // Merge Cells
            ws.get_Range("D6", "F6").Font.Bold = true; 

            // Table headers and style. 
            printCell(7, 2, "Brand"); 
            printCell(7, 3, "BOM Review"); 
            printCell(7, 4, "BOM"); 
            printCell(7, 5, "Router"); 
            printCell(7, 6, "PIR");  
            ws.get_Range("B6", "F7").Font.Bold = true;   

            // Brand rows. 
            printCell(8, 2, "AAL");
            printCell(9, 2, "Beacon");
            printCell(10, 2, "CBurg/Canada");
            printCell(11, 2, "Columbia");
            printCell(12, 2, "HCC");
            printCell(13, 2, "HLOL");
            printCell(14, 2, "KIM");
            printCell(15, 2, "Security");
            printCell(16, 2, "Spaulding");
            printCell(17, 2, "Total");
            ws.Cells[17, 2].Font.Bold = true;
        }

        /// <summary>
        /// Prints data about the processes. 
        /// </summary>
        private void printProcess()
        {
            printCell(1, 10, "ETO's Waiting to Process");
            ws.Cells[1, 10].Font.Bold = true;
            ws.Cells[1, 10].Font.Underline = true;

            printCell(2, 10, "C'Burg");
            ws.Cells[2, 10].Font.Bold = true;
            ws.Cells[2, 10].Font.Underline = true;

            printCell(5, 10, "Reviewing ETO");
            ws.Cells[5, 10].Font.Bold = true;

            printCell(6, 10, "Beacon");
            ws.Cells[6, 10].Font.Bold = true;
            ws.Cells[6, 10].Font.Underline = true;

            printCell(9, 10, "Kim/AAL/Security");
            ws.Cells[9, 10].Font.Bold = true;
            ws.Cells[9, 10].Font.Underline = true;

            printCell(12, 10, "Working on BOM's");
            ws.Cells[12, 10].Font.Bold = true;
            ws.Cells[12, 10].Font.Underline = true;

            printCell(13, 10, "Beacon");
            ws.Cells[13, 10].Font.Bold = true;
            ws.Cells[13, 10].Font.Underline = true;

            printCell(2, 12, "Brand");
            ws.Cells[2, 12].Font.Bold = true;

            printCell(2, 13, "C'Burg");
            ws.Cells[2, 13].Font.Bold = true;

            printCell(5, 12, "Average Days processing ETO's");      // Also prints calculated average.
            ws.Cells[5, 12].Font.Bold = true;
            printCell(6, 12, averageDuration() + " Days");
        }

        /// <summary>
        /// Prints number in Queue, number with BOMs, and the number that are new each day as well as how many were processed the day before. 
        /// </summary>
        private void printOverall()
        {
            printCell(1, 1, "Count in Que");
            printCell(2, 1, "Count with BOM (SAP)");
            printCell(2, 2, retriever.numHits.ToString());
            printCell(3, 1, "Processed since last report");
            printCell(4, 1, "Count New in Que");
        }

        /// <summary>
        /// This method prints labels for the section of the report that is hidden. 
        /// </summary>
        private void printHiddenColumns()
        {
            printBrandLabels();
            printBrandLabelInfo();
        }

        /// <summary>
        /// This prints and merges the cells for each brand in the hidden section. 
        /// </summary>
        private void printBrandLabels()
        {            
            var range = ws.get_Range("P17", "S17");
            range.Merge();
            range.Value2 = "AAL";
            range.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            range = ws.get_Range("T17", "W17");
            range.Merge();
            range.Value2 = "Beacon";
            range.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            range = ws.get_Range("X17", "AA17");
            range.Merge();
            range.Value2 = "Columbia";
            range.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            range = ws.get_Range("AB17", "AE17");
            range.Merge();
            range.Value2 = "HCC";
            range.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            range = ws.get_Range("AF17", "AI17");
            range.Merge();
            range.Value2 = "HLOL";
            range.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            range = ws.get_Range("AJ17", "AM17");
            range.Merge();
            range.Value2 = "KIM";
            range.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            range = ws.get_Range("AN17", "AQ17");
            range.Merge();
            range.Value2 = "SEC";
            range.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            range = ws.get_Range("AR17", "AU17");
            range.Merge();
            range.Value2 = "SPAU";
            range.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;

            range = ws.get_Range("AV17", "AY17");
            range.Merge();
            range.Value2 = "Cburg/Canada";
            range.Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
        }

        /// <summary>
        /// This prints the steps for the brand information in the hidden section. 
        /// </summary>
        private void printBrandLabelInfo()
        {
            int lineBuf = 16;
            
            for (int i = 0; i < 9; i++ )
            {
                printCell(18, i + lineBuf, "BR");
                lineBuf++;
                printCell(18, i + lineBuf, "BOM");
                lineBuf++;
                printCell(18, i + lineBuf, "Router");
                lineBuf++;
                printCell(18, i + lineBuf, "PIR");
            }
        }

        /// <summary>
        /// This method adds all of the formulas to the excek sheet. 
        /// </summary>
        private void addFormulas()
        {
            ws.Cells[1, 2].Formula = "=COUNTIF(N:N,\"0\")";
            // ws.Cells[2, 2].Formula = "=COUNTIF(D:D,\"yes\")";
            ws.Cells[4, 2].Formula = "=COUNTIF(H:H,\"07/14/17\")";

            // Check to see if today is a Monday
            int daysBack = (DateTime.Today.DayOfWeek.ToString() == "Monday") ? 3 : 1;

            // New in Queue
            ws.Cells[4, 2].Formula = "=COUNTIF(H:H, " + "\"" + DateTime.Today.AddDays(-daysBack).ToString("MM/dd/yyyy") + "\")";

            // BOM Review
            ws.Cells[8, 3].Formula = "=COUNTIF(P:P,\"1\")";
            ws.Cells[9, 3].Formula = "=COUNTIF(T:T,\"1\")";
            ws.Cells[10, 3].Formula = "=COUNTIF(AV:AV,\"1\")";
            ws.Cells[11, 3].Formula = "=COUNTIF(X:X,\"1\")";
            ws.Cells[12, 3].Formula = "=COUNTIF(AB:AB,\"1\")";
            ws.Cells[13, 3].Formula = "=COUNTIF(AF:AF,\"1\")";
            ws.Cells[14, 3].Formula = "=COUNTIF(AJ:AJ,\"1\")";
            ws.Cells[15, 3].Formula = "=COUNTIF(AN:AN,\"1\")";
            ws.Cells[16, 3].Formula = "=COUNTIF(AR:AR, \"1\")";
            ws.Cells[17, 3].Formula = "=SUM(C8:C16)";

            // BOM
            ws.Cells[8, 4].Formula = "=COUNTIF(Q:Q,\"1\")";
            ws.Cells[9, 4].Formula = "=COUNTIF(U:U,\"1\")";
            ws.Cells[10, 4].Formula = "=COUNTIF(AW:AW,\"1\")";
            ws.Cells[11, 4].Formula = "=COUNTIF(Y:Y,\"1\")";
            ws.Cells[12, 4].Formula = "=COUNTIF(AC:AC,\"1\")";
            ws.Cells[13, 4].Formula = "=COUNTIF(AG:AG,\"1\")";
            ws.Cells[14, 4].Formula = "=COUNTIF(AK:AK,\"1\")";
            ws.Cells[15, 4].Formula = "=COUNTIF(AO:AO,\"1\")";
            ws.Cells[16, 4].Formula = "=COUNTIF(AS:AS,\"1\")";
            ws.Cells[17, 4].Formula = "=SUM(D8:D16)";

            // Router
            ws.Cells[8, 5].Formula = "=COUNTIF(R:R,\"1\")";
            ws.Cells[9, 5].Formula = "=COUNTIF(V:V,\"1\")";
            ws.Cells[10, 5].Formula = "=COUNTIF(AX:AX,\"1\")";
            ws.Cells[11, 5].Formula = "=COUNTIF(Z:Z,\"1\")";
            ws.Cells[12, 5].Formula = "=COUNTIF(AD:AD,\"1\")";
            ws.Cells[13, 5].Formula = "=COUNTIF(AH:AH,\"1\")";
            ws.Cells[14, 5].Formula = "=COUNTIF(AL:AL,\"1\")";
            ws.Cells[15, 5].Formula = "=COUNTIF(AP:AP,\"1\")";
            ws.Cells[16, 5].Formula = "=COUNTIF(AT:AT,\"1\")";
            ws.Cells[17, 5].Formula = "=SUM(E8:E16)";

            // PIR
            ws.Cells[8, 6].Formula = "=COUNTIF(S:S,\"1\")";
            ws.Cells[9, 6].Formula = "=COUNTIF(W:W,\"1\")";
            ws.Cells[10, 6].Formula = "=COUNTIF(AY:AY,\"1\")";
            ws.Cells[11, 6].Formula = "=COUNTIF(AA:AA,\"1\")";
            ws.Cells[12, 6].Formula = "=COUNTIF(AE:AE,\"1\")";
            ws.Cells[13, 6].Formula = "=COUNTIF(AI:AI,\"1\")";
            ws.Cells[14, 6].Formula = "=COUNTIF(AM:AM,\"1\")";
            ws.Cells[15, 6].Formula = "=COUNTIF(AQ:AQ,\"1\")";
            ws.Cells[16, 6].Formula = "=COUNTIF(AU:AU,\"1\")";
            ws.Cells[17, 6].Formula = "=SUM(F8:F16)";

            //Total
            ws.Cells[17, 7].Formula = "=SUM(C17:F17)";

            // Hidden Rows 
            ws.Cells[19, 14].Formula = "=IF(COUNTIF($B$19:$B19, $B19) > 1, 1, 0)";
            ws.Cells[19, 15].Formula = "=RIGHT(LEFT(B19, 7), 4)";
            ws.Cells[19, 16].Formula = "=IF(AND($O19=\"AALC\", $F19=\"BP1\", $N19=0), 1, 0)"; 
            ws.Cells[19, 17].Formula = "=IF(AND($O19=\"AALC\", $F19=13, $N19=0), 1, 0)";
            ws.Cells[19, 18].Formula = "=IF(AND($O19=\"AALC\", $F19=\"ETO\", $N19=0), 1, 0)";
            ws.Cells[19, 19].Formula = "=IF(AND($O19=\"AALC\", $F19=\"PIR\", $N19=0), 1, 0)";
            ws.Cells[19, 20].Formula = "=IF(AND($O19=\"BEAC\", $F19=\"BP1\", $N19=0), 1, 0)"; // T
            ws.Cells[19, 21].Formula = "=IF(AND($O19=\"BEAC\", $F19=\"BP2\", $N19=0), 1, 0)";
            ws.Cells[19, 22].Formula = "=IF(AND($O19=\"BEAC\", $F19=\"BP3\", $N19=0), 1, 0)";
            ws.Cells[19, 23].Formula = "=IF(AND($O19=\"BEAC\", $F19=\"PIR\", $N19=0), 1, 0)";
            ws.Cells[19, 24].Formula = "=IF(AND($O19=\"COLC\", $F19=\"BP1\", $N19=0), 1, 0)";
            ws.Cells[19, 25].Formula = "=IF(AND($O19=\"COLC\", $F19=\"BP2\", $N19=0), 1, 0)";
            ws.Cells[19, 26].Formula = "=IF(AND($O19=\"COLC\", $F19=\"ETO\", $N19=0), 1, 0)";
            ws.Cells[19, 27].Formula = "=IF(AND($O19=\"COLC\", $F19=\"PIR\", $N19=0), 1, 0)";
            ws.Cells[19, 28].Formula = "=IF(AND($O19=\"HCCB\", $F19=\"BP1\", $N19=0), 1, 0)";
            ws.Cells[19, 29].Formula = "=IF(AND($O19=\"HCCB\", $F19=\"BP2\", $N19=0), 1, 0)";
            ws.Cells[19, 30].Formula = "=IF(AND($O19=\"HCCB\", $F19=\"ETO\", $N19=0), 1, 0)";
            ws.Cells[19, 31].Formula = "=IF(AND($O19=\"HCCB\", $F19=\"PIR\", $N19=0), 1, 0)";
            ws.Cells[19, 32].Formula = "=IF(AND(OR($O19=\"HLOL\", $O19=\"HSLS\"), $F19=\"BP1\", $N19=0), 1, 0)";
            ws.Cells[19, 33].Formula = "=IF(AND(OR($O19=\"HLOL\", $O19=\"HSLS\"), $F19=\"BP2\", $N19=0), 1, 0)";
            ws.Cells[19, 34].Formula = "=IF(AND(OR($O19=\"HLOL\", $O19=\"HSLS\"), $F19=\"ETO\", $N19=0), 1, 0)";
            ws.Cells[19, 35].Formula = "=IF(AND(OR($O19=\"HLOL\", $O19=\"HSLS\"), $F19=\"PIR\", $N19=0), 1, 0)";
            ws.Cells[19, 36].Formula = "=IF(AND($O19=\"KIMC\", $F19=\"BP1\", $N19=0), 1, 0)";
            ws.Cells[19, 37].Formula = "=IF(AND($O19=\"KIMC\", $F19=13, $N19=0), 1, 0)";
            ws.Cells[19, 38].Formula = "=IF(AND($O19=\"KIMC\", $F19=\"ETO\", $N19=0), 1, 0)";
            ws.Cells[19, 39].Formula = "=IF(AND($O19=\"KIMC\", $F19=\"PIR\", $N19=0), 1, 0)";
            ws.Cells[19, 40].Formula = "=IF(AND($O19=\"SECC\", $F19=\"BP1\", $N19=0), 1, 0)"; 
            ws.Cells[19, 41].Formula = "=IF(AND($O19=\"SECC\", $F19=\"BP2\", $N19=0), 1, 0)";
            ws.Cells[19, 42].Formula = "=IF(AND($O19=\"SECC\", $F19=\"ETO\", $N19=0), 1, 0)";
            ws.Cells[19, 43].Formula = "=IF(AND($O19=\"SECC\", $F19=\"PIR\", $N19=0), 1, 0)";
            ws.Cells[19, 44].Formula = "=IF(AND($O19=\"SPAU\", $F19=\"BP1\", $N19=0), 1, 0)";
            ws.Cells[19, 45].Formula = "=IF(AND($O19=\"SPAU\", $F19=\"BP2\", $N19=0), 1, 0)";
            ws.Cells[19, 46].Formula = "=IF(AND($O19=\"SPAU\", $F19=\"ETO\", $N19=0), 1, 0)";
            ws.Cells[19, 47].Formula = "=IF(AND($O19=\"SPAU\", $F19=\"PIR\", $N19=0), 1, 0)";
            ws.Cells[19, 48].Formula = "=IF(AND(OR($F19=19, AND($O19=\"BEAC\", $F19=13)), $N19=0), 1, 0)";
            ws.Cells[19, 49].Formula = "=IF(AND(OR($O19=\"COLC\", $O19=\"HCCB\", $O19=\"HLOL\", $O19=\"SECC\", $O19=\"SPAU\"), $F19=13, $N19=0), 1, 0)";
            ws.Cells[19, 50].Formula = "=IF(AND(OR($O19=\"AALC\", $O19=\"COLC\", $O19=\"HCCB\", $O19=\"HLOL\", $O19=\"SECC\", $O19=\"SPAU\", $O19=\"BEAC\"), $F19=\"BP4\", $N19=0), 1, 0)";
        }

        /// <summary>
        /// This method colors portions of the sheet. 
        /// </summary>
        private void addColor()
        {
            var range = ws.get_Range("B6", "G17");
            range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.PaleGreen);

            range = ws.get_Range("B10", "G10");
            range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);

            range = ws.get_Range("E8", "E17");
            range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);

            range = ws.get_Range("F17");
            range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);

            range = ws.get_Range("J1");
            range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSalmon);
        }

        /// <summary>
        /// This method adds the graph to the excel sheet and formats it. 
        /// </summary>
        private void addGraph()
        {
            const string topLeft = "L2";
            const string bottomRight = "M3";
            const string graphTitle = "ETO Queue Snapshot";

            // Create instance of a chart object. 
            var charts = ws.ChartObjects();
            var chartObject = charts.Add(3000, 20, 300, 300);
            var chart = chartObject.Chart;
            var range = ws.get_Range(topLeft, bottomRight);

            // Setting the details of the chart. 
            chart.ChartType = XlChartType.xlColumnClustered;
            chart.ChartWizard(Source: range, Title: graphTitle);
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
                sumDuration = sumDuration + salesList[row - 19].daysInQueue;
                numOrders++;                                                            // Keeps track of the number of orders. 

                numRow = row;
            }

            return numRow;
        }

        /// <summary>
        /// This method temporarily prints the materials so that they can be copied and pasted into SAP
        /// for finding the information for the BOMs column. 
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
            retriever = new Retreiver(salesList, excel, wbs, wb, ws);
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
            printCell(18, 9, "Started");
            printCell(18, 10, "Additional Notes");
            printCell(18, 11, "Engineering Notes");
            printCell(18, 12, "Ranking 1-5");
            printCell(18, 13, "Days in Queue");
        }

        /// <summary>
        /// This method formats the alignment of the numbers throughout the sheet. 
        /// </summary>
        private void formatAlignment()
        {
            var range = ws.get_Range("A1", "BG17");
            range.HorizontalAlignment = XlHAlign.xlHAlignLeft;
        }

        /// <summary>
        /// This method divides the sum of the order days by the number of orders to find the average days in queue. 
        /// </summary>
        /// <returns> Average days in queue. </returns>
        private double averageDuration()
        {
            return (sumDuration / numOrders);
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

        /// <summary>
        /// This method prints data in a cell. 
        /// </summary>
        /// <param name="i"> The "y" coordinate on a plane. </param>
        /// <param name="j"> The "x" coordinate on a plane. </param>
        /// <param name="value"> The data to be printed in the cell. </param>
        private void printCell(int i, int j, string value)
        {
            ws.Cells[i, j].Value2 = value;             
        }
    }
}
