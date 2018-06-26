using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Wordgenerator.Models.DAL.Films
{
    public class Film
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Number { get; set; }
        public string OwnerAndYear { get; set; }
        public string Country { get; set; }
        public string DurationTime { get; set; }
        public string Language { get; set; }
        public DateTime AgreeementDate { get; set; }
        public string MainCities { get; set; }
        public string Odessa { get; set; }
        public string OtherCities { get; set; }
    }
}