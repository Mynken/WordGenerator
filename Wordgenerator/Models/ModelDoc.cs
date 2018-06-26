using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Wordgenerator.Models
{
    public class ModelDoc
    {
        public int KontrahentId { get; set; }
        public int FilmId { get; set; }
        public int? ThirdAgreementNumber { get; set; }
        public DateTime FilmAgreeementDate { get; set; }
        public string GeneralAgreementType { get; set; }
        public DateTime GeneralAgreementDate { get; set; }
        public string DuplicatedLanguage { get; set; }
        public string City { get; set; }
        public string CinemaName { get; set; }
        public DateTime DemonstrationPeriodFrom { get; set; }
        public DateTime DemonstrationPeriodTo { get; set; }
        public string FilmFormat { get; set; }
        public List<SessionModel> SessionModel { get; set; }
        public string TypeOfFilm { get; set; }
        public string CartoonFilmInfo { get; set; }
        public int FilmPeriodInfo { get; set; }
        public int TimeZoneOffset { get; set; }
        public bool IsPdf { get; set; }
        public string DaysInfo { get; set; }
    }
}