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
    /// TODO: Make more efficient, get rid of all the seperate for loops. 
    /// 
    /// @author: Alexander James Bochel
    /// @version: 6/23/2017
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
        /// <returns> Returns true if exits correctly. </returns>
        public void execute()
        {
            deleteTJ();
            deleteFifteens();
            deleteKVTens();
            
            sortMaterial();

            for (int i = 0; i < salesList.Count; i++)
            {
                convertDate(i);
                findQueueDuration(i);
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
        
        /// <summary>
        /// This method deletes all MSPS's with value of 15 except for Kurt Versen brand. 
        /// </summary>
        public void deleteFifteens()
        {
            for (int i = 0; i < salesList.Count; i++ )
            {
                if (salesList[i].MSPS == "15" && !salesList[i].material.Contains("KV"))
                {
                    salesList.Remove(salesList[i]);
                    i--;
                }
            }
        }

        /// <summary>
        /// This method deletes all KV products with values of 10. 
        /// </summary>
        public void deleteKVTens()
        {
            for (int i = 0; i < salesList.Count; i++)
            {
                if (salesList[i].MSPS == "10" && salesList[i].material.Contains("KV"))
                {
                    salesList.Remove(salesList[i]);
                    i--;
                }
            }
        }

        /// <summary>
        /// This method deletes all sales with a 'TJ' (or just J) material. 
        /// </summary>
        public void deleteTJ()
        {
            for (int i = 0; i < salesList.Count; i++)
            {
                if (salesList[i].material.Contains("J"))
                {
                    salesList.Remove(salesList[i]);
                    i--;
                }
            }
        }

        /// <summary>
        /// This method iterates through every Sales object and retrieves a correctly formatted date. 
        /// </summary>
        /// <param name="i"> Count for the salesList. </param>
        public void convertDate(int i)
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

        /// <summary>
        /// This method iterates through every Sales object and finds the number of days an ETO is 
        /// in the queue for. 
        /// </summary>
        /// <param name="i"> Count for the salesList. </param>
        public void findQueueDuration(int i)
        {
                salesList[i].findQueueDuration();
        }
    }
}
