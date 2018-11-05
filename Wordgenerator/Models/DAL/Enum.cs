using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Web;

namespace Wordgenerator.Models.DAL
{
    public enum Gender
    {
        None,
        Female,
        Male
    }

    public enum Dictionary
    {
        FilmType = 1000,
        FilmFormat = 1001,
        DuplicationLanguage = 1002,
        CartoonFilmInfo = 1003
    }


    public enum TypeOfOrganization
    {
        [Description("TOB")]
        LLC = 2000,
        [Description("FOP")]
        IE = 2001
    }
}