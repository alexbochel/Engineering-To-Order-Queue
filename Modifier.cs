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
    /// Author: Alexander James Bochel
    /// Date Updated: 6/16/2017
    /// 
    /// </summary>
    public class Modifier
    {
        public List<Sales> salesList;

        /// <summary>
        /// This is the constructor for the Modifier object. 
        /// </summary>
        /// <param name="sales"> Takes the list of all of the sales objects created. </param>
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
            sortMaterial();
            convertDate();
            findQueueDuration();
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
        /// This method deletes all MSPS's with value of 15. 
        /// </summary>
        public void deleteFifteens()
        {
            for (int i = 0; i < salesList.Count; i++ )
            {
                if (salesList[i].MSPS == "15")
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
        public void convertDate()
        {
            for (int i = 0; i < salesList.Count; i++)
            {
                salesList[i].doubleDate = double.Parse(salesList[i].date);
                salesList[i].formatDate = DateTime.FromOADate(salesList[i].doubleDate);
            }
        }

        /// <summary>
        /// This method iterates through every Sales object and finds the number of days an ETO is 
        /// in the queue for. 
        /// </summary>
        public void findQueueDuration()
        {
            for (int i = 0; i < salesList.Count; i++)
            {
                salesList[i].findQueueDuration();
            }
        }
    }
}
