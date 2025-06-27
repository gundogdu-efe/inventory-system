using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace EnvanterSistemi.ViewModel
{
    public class Listeler
    {
        public IEnumerable<SelectListItem> Kategoriler { get; set; }
        public IEnumerable<SelectListItem> EnvanterTur { get; set; }
        
        public IEnumerable<SelectListItem> Birimler { get; set; }

        public IEnumerable<SelectListItem>MarkaModel { get; set; }

    }
}