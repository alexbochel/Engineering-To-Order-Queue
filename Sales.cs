using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1 
{
    /// <summary>
    /// 
    /// This class contains information about each individual row on the excel sheet. 
    /// 
    /// @author: Alexander James Bochel
    /// @version: 7/31/2017
    /// 
    /// </summary>
    public class Sales
    {
        // Each variable is a cell in the row for each sales order in excel. 
        private string _salesNum;
        private string _material;
        private string _description;
        private string _BOM;
        private string _MSPS;
        private string _MRPC;
        private string _quantity;
        private string _date;            
        private DateTime _formatDate;    
        private double _doubleDate;
        private int _daysInQueue;
        private DateTime todaysDate;

        /// <summary>
        /// Getter/Setter: Sales number for the part. 
        /// </summary>
        public string salesNum
        {
            get { return _salesNum; }
            set { _salesNum = value; }
        }

        /// <summary>
        /// Getter/Setter: Material used in part production
        /// </summary>
        public string material
        {
            get { return _material; }
            set { _material = value; }
        }
        /// <summary>
        /// Getter/Setter: Description of the part. 
        /// </summary>
        public string description
        {
            get { return _description; }
            set { _description = value; }
        }
        /// <summary>
        /// Getter/Setter: Whether or not the part has a BOM.  
        /// </summary>
        public string BOM
        {
            get { return _BOM; }
            set { _BOM = value; }
        }
        /// <summary>
        /// Getter/Setter: MSPS Number for the part. 
        /// </summary>
        public string MSPS
        {
            get { return _MSPS; }
            set { _MSPS = value; }
        }
        /// <summary>
        /// Getter/Setter: MRPC value for the part. 
        /// </summary>
        public string MRPC
        {
            get { return _MRPC; }
            set { _MRPC = value; }
        }
        /// <summary>
        /// Getter/Setter: How many of the part must be created. 
        /// </summary>
        public string quantity
        {
            get { return _quantity; }
            set { _quantity = value; }
        }
        /// <summary>
        /// Getter/Setter: The date in which the order was added to the queue. 
        /// </summary>
        public string date
        {
            get { return _date; }
            set { _date = value; }
        }
        /// <summary>
        /// Getter/Setter: DateTime version of the parts creation date. 
        /// </summary>
        public DateTime formatDate
        {
            get { return _formatDate; }
            set { _formatDate = value; }
        }

        /// <summary>
        /// Getter/Setter: The double version of the parts creation date (this is the weird number you get from Excel).
        /// </summary>
        public double doubleDate
        {
            get { return _doubleDate; }
            set { _doubleDate = value; }
        }
        /// <summary>
        /// Getter/Setter: The amount of days the part has been in the queue. 
        /// </summary>
        public int daysInQueue
        {
            get { return _daysInQueue; }
            set { _daysInQueue = value; }
        } 

        /// <summary>
        /// This constructor sets the default BOM status to "no". 
        /// </summary>
        public Sales()
        {
            BOM = "no";   
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
