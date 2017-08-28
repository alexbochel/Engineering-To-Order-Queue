using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Windows.Forms;

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
        static string fileExcel;

        [STAThread]
        static void Main(string[] args)
        {            
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "Open Previous ETO Report";
            dialog.Filter = "Excel Files|*.xls;*.xlsx";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                fileExcel = dialog.FileName;
            }
            
            openAndExecute();
            print();
        }

        private static void openAndExecute()
        {
            retreiver = new Retreiver();
            retreiver.runOREP();
            


            reader = new Reader(fileExcel);
            reader.modifier.execute();
        }

        private static void print()
        {
            Printer printer = new Printer(reader.modifier.salesList,
                reader.excel, reader.wbs, reader.wb, reader.ws);
        }
    }
}