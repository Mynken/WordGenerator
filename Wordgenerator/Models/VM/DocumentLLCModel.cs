using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Wordgenerator.Models.DAL.Films;
using Wordgenerator.Models.DAL.Kontrahent;

namespace Wordgenerator.Models.VM
{
    public class DocumentLLCModel
    {
        public KontrahentLLC Kontrahent { get; set; }
        public Film Film { get; set; }
        public List<Trailer> Trailers { get; set; }
        public ModelDoc DataForDoc { get; set; }
    }
}