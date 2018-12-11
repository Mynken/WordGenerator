using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Wordgenerator.Models.DAL.Kontrahent;

namespace Wordgenerator.Models.DAL.Additional
{
    public class AdditionalLLC
    {
        public string FirstAgreementNumber { get; set; }
        public string GeneralAgreementType { get; set; }
        public DateTime GeneralAgreementDate { get; set; }
        public DateTime CityDate { get; set; }
        public KontrahentLLC Kontrahent { get; set; }
        public string CartoonFilmInfo { get; set; }
        public int TimeOffset { get; set; }
    }
}