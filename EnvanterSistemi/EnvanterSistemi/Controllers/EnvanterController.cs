using EnvanterSistemi.Models;
using EnvanterSistemi.ViewModel;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Helpers;
using System.Web.Mvc;
using System.Xml.Linq;

namespace EnvanterSistemi.Controllers
{
    [Authorize]
    public class EnvanterController : Controller
    {
        string cnnstr = "data source=.\\SQLEXPRESS;initial catalog=EnvanterSistemi;integrated security=SSPI";
        EnvanterSistemiEntities db = new EnvanterSistemiEntities();
        Listeler lst = new Listeler();
        // GET: Envanter
        
        public ActionResult Index()
        {

            var model = db.Envanters.ToList();
            return View(model);
        }
        [HttpGet]
        public ActionResult Ekle()
        {
            lst.Kategoriler = new SelectList(db.VarlikKategoris, "id", "kategoriAdi");
            lst.EnvanterTur = new SelectList(db.EnvanterTür, "id", "turAdi");
            lst.Birimler = new SelectList(db.Birimlers, "birimAdi", "birimAdi");
            return View(lst);
        }
        [HttpGet]
        public ActionResult Guncelle(int? id)
        {
            var model = db.Envanters.Where(x => x.id == id).FirstOrDefault();

            ViewBag.envanterid = id;

            string kategoriid = Convert.ToString(model.kategoriId);
            

            List<SelectListItem> kategori = (from p in db.VarlikKategoris.AsEnumerable()
                                              select new SelectListItem
                                              {
                                                  Text = p.kategoriAdi,
                                                  Value = p.id.ToString()
                                              }).ToList();
            kategori.Find(c => c.Value == kategoriid).Selected = true;

            ViewBag.kategoriler = kategori;

            string turid = Convert.ToString(model.turId);

            List<SelectListItem> tur = (from p in db.EnvanterTür.AsEnumerable()
                                             select new SelectListItem
                                             {
                                                 Text = p.turAdi,
                                                 Value = p.id.ToString()
                                             }).ToList();
            tur.Find(c => c.Value == turid).Selected = true;

            ViewBag.turler = tur;

            string birim =model.ilgiliBirim;


            List<SelectListItem> birimler = (from p in db.Birimlers.AsEnumerable()
                                             select new SelectListItem
                                             {
                                                 Text = p.birimAdi,
                                                 Value = p.birimAdi
                                             }).ToList();
            birimler.Find(c => c.Text == birim).Selected = true;

            ViewBag.Birimler = birimler;




            ViewBag.kodu = model.kodu;
            ViewBag.aciklama = model.aciklama;
            return View();
        }
        [HttpPost]
        public ActionResult Guncelle(FormCollection p)
        {
            var envanterid = p["envanterid"];
            string ilgiliBirim = p["DrpBirim"];
            string kategori = p["DrpKategori"];
            string tur = p["DrpTur"];
            string aciklama = p["aciklama"];

            using (SqlConnection connect = new SqlConnection(cnnstr)) {
                connect.Open();
                string guncel = "update [EnvanterSistemi].[dbo].[Envanter] set ilgiliBirim='"+ilgiliBirim+"',kategoriId='"+kategori+"',turId='"+tur+"',aciklama='"+aciklama+"' where id='"+envanterid+"'";
                SqlCommand guncelcmd = new SqlCommand(guncel, connect);
                guncelcmd.ExecuteNonQuery();
                connect.Close();
            }
            TempData["Mesaj"] = envanterid + " Envanter kodu başarılı bir şekilde güncellenmiştir.";



            return RedirectToAction("Index","Envanter");
        }
        [HttpPost]
        public ActionResult Ekle(FormCollection p)
        {
            string kategori = p["DrpKategori"];
            string tur = p["DrpTur"];
            string kod = p["txtKodu"];
            string ilgiliBirim = p["ilgiliBirim"];
            string aciklama = p["aciklama"];
            string username = p["username"];
            string kayitTarihi = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");


            using (SqlConnection connect = new SqlConnection(cnnstr))
            {
                connect.Open();
                string envanterekle = "insert into [EnvanterSistemi].[dbo].[Envanter] (kategoriId,turId,kodu,ilgiliBirim,aciklama,status,zstatus,kaydedenPersonel,kayitTarihi) values ('" + kategori + "','" + tur + "','" + kod + "','" + ilgiliBirim + "','" + aciklama + "','True','False','" + username + "','" + kayitTarihi + "')";
                SqlCommand envcmd = new SqlCommand(envanterekle, connect);
                envcmd.ExecuteNonQuery();

                string ortakk = "select SUBSTRING('"+kod+"',0,4)";
                SqlCommand ortcmd = new SqlCommand(ortakk, connect);
                ortcmd.ExecuteNonQuery();
                string ortakkod = Convert.ToString(ortcmd.ExecuteScalar());

                string siradakii = "select RIGHT('"+kod+"',CHARINDEX('0',REVERSE('"+kod+"'))-1)";
                SqlCommand sircmd = new SqlCommand(siradakii, connect);
                sircmd.ExecuteNonQuery();
                int siradakino = Convert.ToInt32(sircmd.ExecuteScalar());


                string siradakix = "update [EnvanterSistemi].[dbo].[TurSiraNo] set sonkullanilanNo='"+siradakino+"' where kategoriKodu='"+ortakkod+"'";
                SqlCommand siracmdx = new SqlCommand(siradakix, connect);
                siracmdx.ExecuteNonQuery();


                connect.Close();                                                                                                                                                                                                                 

            }
            TempData["Mesaj"] = kod+" Envanter kodu başarılı bir şekilde kaydedilmiştir.";
            return RedirectToAction("Index", "Envanter");
        }
        [HttpPost]
        public JsonResult TurGetir(int? p)
        {

            int ID = Convert.ToInt32(p);
            var turler = (from x in db.EnvanterTür
                          join y in db.VarlikKategoris on x.kategoriId equals y.id
                          where x.kategoriId == ID
                          select new
                          {
                              Text = x.turAdi,
                              Value = x.id
                          }).ToList();

            return Json(turler, JsonRequestBehavior.AllowGet);
        }

        public JsonResult KodGetir(string p)
        {
            string yyenikod = "";
            string cnnstr = "data source=.\\SQLEXPRESS;initial catalog=EnvanterSistemi;integrated security=SSPI";
            int turid = Convert.ToInt32(p);
            using (SqlConnection connect = new SqlConnection(cnnstr))
            {
                connect.Open();
                string kodlist1 = "select ortakKod FROM [EnvanterSistemi].[dbo].[EnvanterTür] WHERE id='" + turid + "'";
                SqlCommand kodlist1cmd = new SqlCommand(kodlist1, connect);
                kodlist1cmd.ExecuteNonQuery();
                string ortakkod = Convert.ToString(kodlist1cmd.ExecuteScalar());

                string kodlist2 = "select kategoriKodu FROM [EnvanterSistemi].[dbo].[EnvanterTür] WHERE id='" + turid + "'";
                SqlCommand kodlist2cmd = new SqlCommand(kodlist2, connect);
                kodlist1cmd.ExecuteNonQuery();
                string kategorikodu = Convert.ToString(kodlist2cmd.ExecuteScalar());

                string kodlist3 = "Select Len('" + kategorikodu + "')";
                SqlCommand kodlist3cmd = new SqlCommand(kodlist3, connect);
                kodlist3cmd.ExecuteNonQuery();
                int yenikatbsm = Convert.ToInt32(kodlist3cmd.ExecuteScalar());

                string list = "select sonkullanilanNo FROM [EnvanterSistemi].[dbo].[TurSiraNo] WHERE kategorikodu='" + ortakkod + "'";
                SqlCommand listcmd = new SqlCommand(list, connect);
                listcmd.ExecuteNonQuery();
                int sonkullanilanno = Convert.ToInt32(listcmd.ExecuteScalar());

                int siradakino = sonkullanilanno + 1;

                string kodlist4 = "Select Len('" + siradakino + "')";
                SqlCommand kodlist4cmd = new SqlCommand(kodlist4, connect);
                kodlist4cmd.ExecuteNonQuery();
                int kodsayibsm = Convert.ToInt32(kodlist4cmd.ExecuteScalar());

                int yenisifiradt = 8 - yenikatbsm - kodsayibsm;

                

                if (yenisifiradt == 0)
                {
                    string envlist10 = "select '" + kategorikodu + "'+'" + siradakino + "'";
                    SqlCommand envlist10s = new SqlCommand(envlist10, connect);
                    envlist10s.ExecuteNonQuery();
                    yyenikod = Convert.ToString(envlist10s.ExecuteScalar());
                }
                else if (yenisifiradt == 1)
                {
                    string envlist10 = "select '" + kategorikodu + "'+'0'+'" + siradakino + "'";
                    SqlCommand envlist10s = new SqlCommand(envlist10, connect);
                    envlist10s.ExecuteNonQuery();
                    yyenikod = Convert.ToString(envlist10s.ExecuteScalar());
                }
                else if (yenisifiradt == 2)
                {
                    string envlist10 = "select '" + kategorikodu + "'+'00'+'" + siradakino + "'";
                    SqlCommand envlist10s = new SqlCommand(envlist10, connect);
                    envlist10s.ExecuteNonQuery();
                    yyenikod = Convert.ToString(envlist10s.ExecuteScalar());
                }
                else if (yenisifiradt == 3)
                {
                    string envlist10 = "select '" + kategorikodu + "'+'000'+'" + siradakino + "'";
                    SqlCommand envlist10s = new SqlCommand(envlist10, connect);
                    envlist10s.ExecuteNonQuery();
                    yyenikod = Convert.ToString(envlist10s.ExecuteScalar());
                }
                connect.Close();
            }
            
            return Json(yyenikod, JsonRequestBehavior.AllowGet);
        }
    }
}