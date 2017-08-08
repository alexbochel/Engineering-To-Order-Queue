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
    /// 
    /// This class modifies an array of all of the sales in order to manipulate the order and contents
    /// which will then print out onto the excel document. 
    /// 
    /// @author: Alexander James Bochel
    /// @version: 7/21/2017
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
                
                this.findQueueDuration(j);

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
        public void sortMaterial()
        {
            salesList = salesList.OrderBy(o => o.material).ToList();
        }

        /// <summary>
        /// Deletes entries that have been in the queue for over 100 days. 
        /// </summary>
        /// <param name="i"> Index in the list. </param>
        /// <param name="deleted"> Whether or not an item has been deleted this iteration. </param>
        private void tweaks(ref int i, ref bool deleted)
        {
            if (!deleted && salesList[i].daysInQueue > 100)
            {
                remove(ref i, ref deleted);
            }
        }

        /// <summary>
        /// Deletes all Prescott ETOs from the list. 
        /// </summary>
        /// <param name="i"> Index in the list. </param>
        /// <param name="deleted"> Whether or not an item has been deleted this iteration. </param>
        private void deletePres(ref int i, ref bool deleted)
        {
            if (!deleted && salesList[i].material.Contains("PRES"))
            {
                remove(ref i, ref deleted);
            }
        }

        /// <summary>
        /// Deletes alll status 15 ETOs from the list.  
        /// </summary>
        /// <param name="i"> Index in the list. </param>
        /// <param name="deleted"> Whether or not an item has been deleted this iteration. </param>
        private void deleteFifteens(ref int i, ref bool deleted)
        {
            if (!deleted && salesList[i].MSPS == "15" && !salesList[i].material.Contains("KV"))
            {
                remove(ref i, ref deleted);
            }
        }

        /// <summary>
        /// Deletes all status 10 KV ETOs from the list. 
        /// </summary>
        /// <param name="i"> Index in the list. </param>
        /// <param name="deleted"> Whether or not an item has been deleted this iteration. </param>
        private void deleteKVTens(ref int i, ref bool deleted)
        {
            if (!deleted && salesList[i].MSPS == "10" && salesList[i].material.Contains("KV"))
            {
                remove(ref i, ref deleted);
            }
        }

        /// <summary>
        /// Deletes all ETOs handled by Tijuana. 
        /// </summary>
        /// <param name="i"> Index in the list. </param>
        /// <param name="deleted"> Whether or not an item has been deleted this iteration. </param>
        private void deleteTJ(ref int i, ref bool deleted)
        {
            if (!deleted && salesList[i].material.Contains("J"))
            {
                remove(ref i, ref deleted);
            }
        }

        /// <summary>
        /// This method removes an entry from the list. 
        /// </summary>
        /// <param name="k"> Index of the item to be removed. </param>
        /// <param name="delete"> Always enters as false, method changes it to true. </param>
        private void remove(ref int k, ref bool delete)
        {
            salesList.Remove(salesList[k]);
            delete = true;
            k--;
        }
        
        /// <summary>
        /// Finds the number of days each ETO has been in the queue and throws an exception if incorrect data is entered. 
        /// </summary>
        /// <param name="i"> Index in the list. </param>
        private void findQueueDuration(int i)
        {
            try
            {
                salesList[i].doubleDate = double.Parse(salesList[i].date);
                salesList[i].formatDate = DateTime.FromOADate(salesList[i].doubleDate);
                salesList[i].findQueueDuration();
            }
            catch
            {
                salesList[i].daysInQueue = 0;
                Console.WriteLine("You Messed Up! Don't forget to have SAP data copied to your clipboard. GO BLUE!");
            }
        }
    }
}

