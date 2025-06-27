using EnvanterSistemi.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace EnvanterSistemi.Controllers
{
    public class MainController : Controller
    {
        EnvanterSistemiEntities db = new EnvanterSistemiEntities();
        // GET: Main
        [Authorize]
        public ActionResult Index()
        {
            var model = db.Envanters.ToList();
            return View(model);
        }
    }
}