using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Wordgenerator.Models
{
    public class SessionModel
    {
        public string NumberOfWeek { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public string SessionInfo { get; set; }
        public DateTime PaymentDate { get; set; }
    }
}