using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1 
{
    /// <summary>
    /// This class contains information about each individual row on the excel sheet. 
    /// 
    /// Author: Alexander James Bochel
    /// Date Updated: 6/16/2017
    /// 
    /// </summary>
    public class Sales
    {
        // Each variable is a cell in the row for each sales order in excel. 
        public String salesNum { get; set; }
        public String material { get; set; }
        public String description { get; set; }
        public String MSPS { get; set; }
        public String MRPC { get; set; }
        public String quantity { get; set; }
        public String date { get; set; }            // String read in by reader
        public DateTime formatDate { get; set; }    // Correctly formatted date to be printed.
        public double doubleDate { get; set; }
        public int daysInQueue { get; set; }
        private DateTime todaysDate;

        public Sales()
        {
            // Nothing to do here.    
        }

        /// <summary>
        /// This method finds how long an ETO has been in the queue. 
        /// </summary>
        public void findQueueDuration()
        {
            todaysDate = DateTime.Today; // Find todays date. 
            daysInQueue = -1 * Convert.ToInt32((formatDate - todaysDate).TotalDays);    // Multiply by negative one to make days number positive.
        }
    }
}
