using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Wordgenerator.Models.DAL.Kontrahent
{
    public class KontrahentBase
    {
        public int Id { get; set; }
        public string Number { get; set; }
        public int TypeOfOrganization { get; set; }
        public string FullName { get; set; }
        public string ActingUnder { get; set; }
        public string Adress { get; set; }
        public string CurrentBankAccount { get; set; }
        public string Mfo { get; set; }
        public string IdentificationCode { get; set; }
        public string RegistrationLicense { get; set; }
        public string TaxInfo { get; set; }
        public string Signature { get; set; }
    }
}