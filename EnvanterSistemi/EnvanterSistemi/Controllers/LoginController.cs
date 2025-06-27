using EnvanterSistemi.Models;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Security;

namespace EnvanterSistemi.Controllers
{
    public class LoginController : Controller
    {
        // GET: Login
        EnvanterSistemiEntities db = new EnvanterSistemiEntities();
        string cnnstr = "data source=.\\SQLEXPRESS;initial catalog=EnvanterSistemi;integrated security=SSPI";
        [HttpGet]
        public ActionResult Index()
        {
            return View();
        }
        [HttpPost]
        public ActionResult Index(Admin p)
        {
            
            var kullanici = db.Admins.Where(x => x.username == p.username && x.password == p.password && x.status == true).FirstOrDefault();
            if (kullanici!=null) 
            {
                FormsAuthentication.SetAuthCookie(p.username,false);
                Session["username"] = p.username;

                return RedirectToAction("Index","Main");
            }

            ViewBag.mesaj = "Kullanıcı veya Şifreyi yanlış girdiniz.";
            return View();
        }
        public ActionResult Logout()
        {
            FormsAuthentication.SignOut();
            return RedirectToAction("Index","Login");
        }
        [HttpGet]
        public ActionResult KayitOl()
        {
            return View(); 
        }
        [HttpPost]
        public ActionResult KayitOl(FormCollection p)
        {
            string email = p["email"];
            string name = p["name"];
            string username = p["username"];
            string password = p["password"];
            string confirmpassword = p["confirmpassword"];

            if (password==confirmpassword) 
            {
                string cnnstr = "data source=DESKTOP-L3P1LOQ\\SQLEXPRESS;initial catalog=EnvanterSistemi;integrated security=True";
                using (SqlConnection connect = new SqlConnection(cnnstr))
                {
                    connect.Open();
                    string adminekle = "insert into [EnvanterSistemi].[dbo].[Admin] (adSoyad,username,email,password,status) values ('" + name + "','" + username + "','" + email + "','" + password + "','True')";
                    SqlCommand admincmd=new SqlCommand(adminekle,connect);
                    admincmd.ExecuteNonQuery();
                    connect.Close();
                }
                ViewBag.mesaj = "Kayıt işleminiz başarılı bir şekilde gerçekleşmiştir.";
                return RedirectToAction("Index","Login");
            }

            ViewBag.mesaj = "Girdiğiniz parolalar uyuşmamaktadır.";
            return View();
        }
        [HttpGet]
        public ActionResult SifremiUnuttum() 
        { 
            return View();
        }
        [HttpPost]
        public ActionResult SifremiUnuttum(Admin p)
        {
            
            using (SqlConnection connect = new SqlConnection(cnnstr))
            {
                connect.Open();

                string email = p.email;
                string username = p.username;
                TempData["email"] = p.email;
                var test = db.Admins.Where(y => y.email == p.email && y.status==true).ToList();
                
                if (test == null)
                {
                    ViewBag.mesaj = "Kullanıcı girişi yapılmadı!";
                    //return RedirectToAction("SifreGir", "Home");
                }
                else if (test != null)
                {
                    //kullanici2 = db.Admins.FirstOrDefault(x=> x.username==username && x.email==email);
                    return RedirectToAction("SifreGir", "Login");
                    
                    //ViewBag.mesaj = "Böyle bir kullanıcı bulunamadı";


                }

                connect.Close();
                
               
            }
            return View();
        }
        [HttpGet]
        public ActionResult SifreGir()
        {
            
            return View(); 
        }
        [HttpPost]
        public ActionResult SifreGir(FormCollection p)
        {
            using (SqlConnection connect = new SqlConnection(cnnstr))
            {
                connect.Open();
                string email = p["email"];
                string password = p["password"];
                string confirmpassword = p["confirmpassword"];

                if (confirmpassword != null && password != null)
                {
                    if(password == confirmpassword)
                    {
                        string query = "update [EnvanterSistemi].[dbo].[Admin] set password='"+password+"' where email='"+email+"'";
                        SqlCommand command2= new SqlCommand(query,connect);
                        command2.ExecuteNonQuery();
                        return RedirectToAction("Index","Login");
                    }
                    else
                    {
                        ViewBag.mesaj = "Girdiğiniz şifreler birbiriyle uyuşmamaktadır!";
                    }
                }
                else
                {
                    ViewBag.mesaj = "Boş alan bıraktınız!";
                }
                connect.Close();

            }
                return View();
        }

    }
}

//viewbag
//redirecttoaction
//firstordefault