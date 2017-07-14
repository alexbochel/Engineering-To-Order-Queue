using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;


namespace ConsoleApplication1
{
    /// <summary>
    /// This class modifies an array of all of the sales in order to manipulate the order and contents
    /// which will then print out onto the excel document. 
    /// 
    /// @author: Alexander James Bochel
    /// @version: 7/12/2017
    /// 
    /// </summary>
    public class Modifier
    {
        /// <summary>
        /// List of sales read from excel sheet. 
        /// </summary>
        public List<Sales> salesList;

        /// <summary>
        /// This is the constructor for the Modifier object. 
        /// </summary>
        /// <param name="list"> Takes the list of all of the sales objects created. </param>
        public Modifier(List<Sales> list)
        {
            salesList = list;
        }

        /// <summary>
        /// This method is called by Main to perform all operations within this class. 
        /// </summary>
        public void execute()
        {
            sortMaterial();
            modifyEntries();
        }

        /// <summary>
        /// This method iterates through the entries and deletes certain items and adds dates to the rest. 
        /// </summary>
        public void modifyEntries()
        {
            // Loops through all of the entries and deletes unneeded ones and finds dates. 
            for (int j = 0; j < salesList.Count; j++)
            {
                bool found = false;
                
                convertDate(j);
                findQueueDuration(j);

                deletePres(ref j, ref found);
                deleteTJ(ref j, ref found);
                deleteKVTens(ref j, ref found);
                deleteFifteens(ref j, ref found);
                tweaks(ref j, ref found);
            }
        }

        /// <summary>
        /// This method sorts the material by alphabetical order using a selection sort implementation. 
        /// </summary>
        /// <returns> Whether or not sort correctly finishes. </returns>
        public void sortMaterial()
        {
            salesList = salesList.OrderBy(o => o.material).ToList();
        }

        private void tweaks(ref int i, ref bool deleted)
        {
            if (!deleted && salesList[i].daysInQueue > 100)
            {
                remove(ref i, ref deleted);
            }
        }

        private void deletePres(ref int i, ref bool deleted)
        {
            if (!deleted && salesList[i].material.Contains("PRES"))
            {
                remove(ref i, ref deleted);
            }
        }
        
        private void deleteFifteens(ref int i, ref bool deleted)
        {
            if (!deleted && salesList[i].MSPS == "15" && !salesList[i].material.Contains("KV"))
            {
                remove(ref i, ref deleted);
            }
        }

        private void deleteKVTens(ref int i, ref bool deleted)
        {
            if (!deleted && salesList[i].MSPS == "10" && salesList[i].material.Contains("KV"))
            {
                remove(ref i, ref deleted);
            }
        }

        private void deleteTJ(ref int i, ref bool deleted)
        {
            if (!deleted && salesList[i].material.Contains("J"))
            {
                remove(ref i, ref deleted);
            }
        }

        private void remove(ref int k, ref bool delete)
        {
            salesList.Remove(salesList[k]);
            delete = true;
            k--;
        }

        private void convertDate(int i)
        {
            try
            {
                salesList[i].doubleDate = double.Parse(salesList[i].date);
                salesList[i].formatDate = DateTime.FromOADate(salesList[i].doubleDate);
            }
            catch
            {
                Console.WriteLine("You Messed Up! Don't forget to have SAP data copied to your clipboard. GO BLUE!");
            }
        }

        private void findQueueDuration(int i)
        {
            salesList[i].findQueueDuration();
        }
    }
}
