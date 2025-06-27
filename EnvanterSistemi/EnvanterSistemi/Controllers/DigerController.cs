using EnvanterSistemi.Models;
using EnvanterSistemi.ViewModel;
using Microsoft.Ajax.Utilities;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;

namespace EnvanterSistemi.Controllers
{
    [Authorize]
    public class DigerController : Controller
    {
        string cnnstr = "data source=.\\SQLEXPRESS;initial catalog=EnvanterSistemi;integrated security=SSPI";
        EnvanterSistemiEntities db = new EnvanterSistemiEntities();
        Listeler lst = new Listeler();

        // GET: Diger


        
        public ActionResult Index()
        {
            var model = db.MarkaModels.ToList();
            return View(model);
        }
        [HttpGet]
        public ActionResult Ekle() 
        {
 
            lst.MarkaModel = new SelectList(db.MarkaModels.Select(y => y.marka).Distinct().ToList(),"marka");

            return View(lst);
        }
        [HttpPost]
        public ActionResult Ekle(FormCollection p)
        {
            


            string exmarka = p["marka"];
            string ymarka = p["ymarka"];         
            string ymodel = p["ymodel"];
  
           

            using (SqlConnection connect = new SqlConnection(cnnstr))
            {
               connect.Open();
                if (exmarka !="")
                {
                    string markaekle = "insert into [EnvanterSistemi].[dbo].[MarkaModel] (marka,model) values ('" + exmarka + "','" + ymodel + "')";
                    SqlCommand markacmd = new SqlCommand(markaekle, connect);
                    markacmd.ExecuteNonQuery();
                }
                else
                {
                    string markaekle = "insert into [EnvanterSistemi].[dbo].[MarkaModel] (marka,model) values ('" + ymarka + "','" + ymodel + "')";
                    SqlCommand markacmd = new SqlCommand(markaekle, connect);
                    markacmd.ExecuteNonQuery();
                }
                    
               
                connect.Close();
            }
            TempData["Mesaj"] = "Marka başarılı bir şekilde eklenmiştir.";



            return RedirectToAction("Index", "Diger");
        }
    }
}