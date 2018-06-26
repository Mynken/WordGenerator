using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Wordgenerator.Models.DAL.Films
{
    public class Trailer
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public int FilmId { get; set; }
    }
}