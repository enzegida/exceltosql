using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToSql.Models
{
    public class CVTSTransaction
    {
       public string BatchCode { get; set; }
       public string TrackingNumber { get; set; }
        public string CardNumber { get; set; }
        public string CardHolderName { get; set; }
        public int AttemptCounter { get; set; } = 0;
        
        public int VoucherAmount { get; set; } = 0;
        public string DueDate { get; set; } = DateTime.Now.ToString();
        public string CurrentStatus { get; set; } = "003";
        public string StatusDate { get; set; } = DateTime.Now.ToString();
        public string CourierCode { get; set; } = "060";
        public string SqDueDate { get; set; } = DateTime.Now.ToString();
        public string Source { get; set; } = "0904003004399";
        public string CustomerNumber { get; set; }
        public string SubTrackFlag { get; set; } = "N";
        public string SubTrackDate { get; set; }
        public string ExpiryDate { get; set; } = "1230";
        public string CourierTrackingNumber { get; set; } = string.Empty;
        public int Seq { get; set; } = 1;

    }
}
