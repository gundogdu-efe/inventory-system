using EnvanterSistemi.Models;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Excel = Microsoft.Office.Interop.Excel;

namespace EnvanterSistemi.Controllers
{
    [Authorize]
    public class ZimmetController : Controller
    {
        EnvanterSistemiEntities db = new EnvanterSistemiEntities();
        string cnnstr = "data source=.\\SQLEXPRESS;initial catalog=EnvanterSistemi;integrated security=SSPI";
        // GET: Zimmet
        
        public ActionResult Index()
        {

            var model = db.Zimmets.ToList();
            return View(model);
           
        }
        public ActionResult Ekle()
        {
            var model = db.Envanters.Where(x=>x.zstatus==false).ToList();
            //personel dropdown
            List<SelectListItem> personellist = new List<SelectListItem>(from x in db.Personels where x.typ==1
                                                                         select new SelectListItem
                                                                         
                                                                         {
                                                                             Value = x.id.ToString(),
                                                                             Text = x.adSoyad
                                                                         }).ToList();

            ViewBag.personel = personellist;

            //birim dropdown

            List<SelectListItem> birimlerlist = new List<SelectListItem>(from x in db.Birimlers
                                                                         select new SelectListItem
                                                                         {
                                                                             Value=x.id.ToString(),
                                                                             Text=x.birimAdi
                                                                         }).ToList();
            ViewBag.birimler = birimlerlist;    

            //odano dropdown
            List<SelectListItem> odaNoList = new List<SelectListItem>(from x in db.Konums
                                                                      select new SelectListItem
                                                                      {
                                                                          Value=x.id.ToString(),
                                                                          Text=x.odaNo
                                                                      }).ToList();
            ViewBag.odaNo = odaNoList;


            return View(model);

        }

        [HttpPost]
        public ActionResult Ekle(FormCollection p)
        {
            

            if (p["kayitID"] == null)
            {
                TempData["Hata"] = " Herhangi bir envanter seçimi yapmadınız!!!!";
            }
            else
            {
                string[] ids = p["kayitID"].Split(new char[] { ',' });
                string zimmetturu = p["zimmetturu"];
                string birimid = p["DrpBirim"];
                string personel = p["DrpPersonel"];
                string lokasyon = p["lokasyon"];
                string oda = p["DrpOda"];
                string znotu = p["aciklama"];
                string username = p["username"];
                string teslimtarihi = p["TeslimTarihi"];
                string kayitTarihi = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                int adet = ids.Length;
                foreach (string id in ids)
                {
                    using (SqlConnection connect = new SqlConnection(cnnstr))
                    {
                        connect.Open();

                        string birim = "SELECT birimAdi FROM[EnvanterSistemi].[dbo].[Birimler] where id ='" + birimid + "' ";
                        SqlCommand birimcmd = new SqlCommand(birim, connect);
                        birimcmd.ExecuteNonQuery();
                        string birimadi = Convert.ToString(birimcmd.ExecuteScalar());

                        string envanter = "SELECT kodu FROM[EnvanterSistemi].[dbo].[Envanter] where id ='" + id + "' ";
                        SqlCommand envantercmd = new SqlCommand(envanter, connect);
                        envantercmd.ExecuteNonQuery();
                        string envanterkodu = Convert.ToString(envantercmd.ExecuteScalar());

                        if (zimmetturu=="Personel")
                        {
                            string zimmetekle = "insert into [EnvanterSistemi].[dbo].[Zimmet] (envanterId,envanterKodu,personelId,birimAdi,teslimTarihi,zimmetTuru,zimmetNotu,lokasyon,odaNo,status,kaydedenPersonel,kayitZamani) values ('" + id + "','" + envanterkodu + "','" + personel + "','" + birimadi + "','" + teslimtarihi + "','" + zimmetturu + "','" + znotu + "','" + lokasyon + "','" + oda + "','True','" + username + "','" + kayitTarihi + "')";
                            SqlCommand zimmetcmd = new SqlCommand(zimmetekle, connect);
                            zimmetcmd.ExecuteNonQuery();
                        }
                        else
                        {
                            string personelx = "SELECT id FROM[EnvanterSistemi].[dbo].[Personel] where adSoyad ='" + birimadi + "' ";
                            SqlCommand perscmd = new SqlCommand(personelx, connect);
                            perscmd.ExecuteNonQuery();
                            int persid = Convert.ToInt32(perscmd.ExecuteScalar());

                            string zimmetekle = "insert into [EnvanterSistemi].[dbo].[Zimmet] (envanterId,envanterKodu,personelId,birimAdi,teslimTarihi,zimmetTuru,zimmetNotu,lokasyon,odaNo,status,kaydedenPersonel,kayitZamani) values ('" + id + "','" + envanterkodu + "','" + persid + "','" + birimadi + "','" + teslimtarihi + "','" + zimmetturu + "','" + znotu + "','" + lokasyon + "','" + oda + "','True','" + username + "','" + kayitTarihi + "')";
                            SqlCommand zimmetcmd = new SqlCommand(zimmetekle, connect);
                            zimmetcmd.ExecuteNonQuery();
                        }
                        

                        
                        

                       
                        string envupt = "update [EnvanterSistemi].[dbo].[Envanter] set zstatus = '1' where id = '"+id+"'";
                        SqlCommand envuptcmd = new SqlCommand(envupt, connect);
                        envuptcmd.ExecuteNonQuery();

                        connect.Close();

                    }
                }
                TempData["Mesaj"] = adet+" adet envanterin zimmet işlemi başarılı bir şekilde gerçekleştirilmiştir.";
            }
            

                return RedirectToAction("Ekle","Zimmet");
        }

        [HttpGet]
        public ActionResult Guncelle(int? id)
        {
            var model = db.Zimmets.Where(x => x.id == id).FirstOrDefault();

            ViewBag.zimmetid = id;
            ViewBag.envanterid = model.envanterId;
            ViewBag.envanterkodu = model.envanterKodu;

            List<SelectListItem> personellist = new List<SelectListItem>(from x in db.Personels
                                                                         where x.typ == 1
                                                                         select new SelectListItem

                                                                         {
                                                                             Value = x.id.ToString(),
                                                                             Text = x.adSoyad
                                                                         }).ToList();

            ViewBag.personel = personellist;

            //birim dropdown

            List<SelectListItem> birimlerlist = new List<SelectListItem>(from x in db.Birimlers
                                                                         select new SelectListItem
                                                                         {
                                                                             Value = x.id.ToString(),
                                                                             Text = x.birimAdi
                                                                         }).ToList();
            ViewBag.birimler = birimlerlist;

            //odano dropdown
            List<SelectListItem> odaNoList = new List<SelectListItem>(from x in db.Konums
                                                                      select new SelectListItem
                                                                      {
                                                                          Value = x.id.ToString(),
                                                                          Text = x.odaNo
                                                                      }).ToList();
            ViewBag.odaNo = odaNoList;

            return View();
        }
        [HttpPost]
        public ActionResult Guncelle(FormCollection p)
        {
            int zimmetid = Convert.ToInt32(p["zimmetid"]);
            

            var envanterid = p["envanterid"];
            string envanterkodu = p["envanterkodu"];
            string zimmetturu = p["zimmetturu"];
            int ilgiliBirim =Convert.ToInt32(p["DrpBirim"]);
            string personel = p["DrpPersonel"];
            string lokasyon = p["lokasyon"];
            string oda = p["DrpOda"];
            string aciklama = p["aciklama"];
            string username = p["username"];
            string teslimtarihi = p["TeslimTarihi"];
            string iadetarihi = p["TeslimTarihi"];
            string kayitTarihi = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            var model = db.Birimlers.Where(x => x.id == ilgiliBirim).FirstOrDefault();

            string birimadi = model.birimAdi;
            //string kategori = p["DrpKategori"];
            //string tur = p["DrpTur"];


            using (SqlConnection connect = new SqlConnection(cnnstr))
            {
                connect.Open();
                if (zimmetturu == "Personel")
                {
                    string zimmetekle = "insert into [EnvanterSistemi].[dbo].[Zimmet] (envanterId,envanterKodu,personelId,birimAdi,teslimTarihi,zimmetTuru,zimmetNotu,lokasyon,odaNo,status,kaydedenPersonel,kayitZamani) values ('" + envanterid + "','" + envanterkodu + "','" + personel + "','" + birimadi + "','" + teslimtarihi + "','" + zimmetturu + "','" + aciklama + "','" + lokasyon + "','" + oda + "','True','" + username + "','" + kayitTarihi + "')";
                    SqlCommand zimmetcmd = new SqlCommand(zimmetekle, connect);
                    zimmetcmd.ExecuteNonQuery();
                }
                else
                {
                    string personelx = "SELECT id FROM[EnvanterSistemi].[dbo].[Personel] where adSoyad ='" + birimadi + "' ";
                    SqlCommand perscmd = new SqlCommand(personelx, connect);
                    perscmd.ExecuteNonQuery();
                    int persid = Convert.ToInt32(perscmd.ExecuteScalar());

                    string zimmetekle = "insert into [EnvanterSistemi].[dbo].[Zimmet] (envanterId,envanterKodu,personelId,birimAdi,teslimTarihi,zimmetTuru,zimmetNotu,lokasyon,odaNo,status,kaydedenPersonel,kayitZamani) values ('" + envanterid + "','" + envanterkodu + "','" + persid + "','" + birimadi + "','" + teslimtarihi + "','" + zimmetturu + "','" + aciklama + "','" + lokasyon + "','" + oda + "','True','" + username + "','" + kayitTarihi + "')";
                    SqlCommand zimmetcmd = new SqlCommand(zimmetekle, connect);
                    zimmetcmd.ExecuteNonQuery();
                }

                
                string guncel = "update [EnvanterSistemi].[dbo].[Zimmet] set iadeTarihi='"+iadetarihi+"',status='false' where id='" + zimmetid + "'";
                SqlCommand guncelcmd = new SqlCommand(guncel, connect);
                guncelcmd.ExecuteNonQuery();

                connect.Close();
            }
            TempData["Mesaj"] = envanterkodu + " Envanter koduna ait zimmet değişikliği başarılı bir şekilde gerçekleştirilmiştir.";



            return RedirectToAction("Index", "Zimmet");

           
        }

        [HttpGet]
        public ActionResult BaskiAl()
        {
            var model = db.Zimmets.ToList();
            return View(model); 
        }
        [HttpPost]
        public ActionResult BaskiAl(FormCollection p)
        {
            if (p["kayitID"] == null)
            {
                TempData["Hata"] = " Herhangi bir envanter seçimi yapmadınız!!!!";
            }
            else
            {
                //Ctrl + k + s kısayolu ile region oluşturulabilir
                #region Form Verileri
                string[] ids = p["kayitID"].Split(new char[] { ',' });
                //string zimmetturu = p["zimmetturu"];
                //string birimid = p["DrpBirim"];
                //string personel = p["DrpPersonel"];
                //string lokasyon = p["lokasyon"];
                //string oda = p["DrpOda"];
                //string znotu = p["aciklama"];
                var firstid = ids[0];
                var modelz = db.Zimmets.Find(int.Parse(firstid));
                var personelid=modelz.personelId;

                var modelp = db.Personels.Find(personelid);

                var adsoyad = modelp.adSoyad;
                var birim = modelp.birimAdi;
                var unvan = modelp.unvan;
                

                string username = p["username"];
                string kayitTarihi = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"); 
                #endregion


                #region Excel Kod Alanı

                Excel.Application application = new Excel.Application();
                Excel.Workbook workbook = application.Workbooks.Add(System.Reflection.Missing.Value);
                Excel.Worksheet worksheet = workbook.ActiveSheet;

                worksheet.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperA4;
                worksheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;

                worksheet.PageSetup.Zoom = 85;
                worksheet.PageSetup.PrintArea = "A1:S32";
                worksheet.PageSetup.BottomMargin = 0.1;
                worksheet.PageSetup.LeftMargin = 0.1;
                worksheet.PageSetup.RightMargin = 0.1;
                worksheet.PageSetup.TopMargin = 0.5;

                worksheet.Columns[1].ColumnWidth = 1;
                worksheet.Columns[2].ColumnWidth = 5;
                worksheet.Columns[3].ColumnWidth = 5;
                worksheet.Columns[4].ColumnWidth = 5;
                worksheet.Columns[5].ColumnWidth = 5;
                worksheet.Columns[6].ColumnWidth = 8.43;
                worksheet.Columns[7].ColumnWidth = 8.43;
                worksheet.Columns[8].ColumnWidth = 18.57;
                worksheet.Columns[9].ColumnWidth = 12;
                worksheet.Columns[10].ColumnWidth = 12;
                worksheet.Columns[11].ColumnWidth = 12;
                worksheet.Columns[12].ColumnWidth = 12;
                worksheet.Columns[13].ColumnWidth = 12;
                worksheet.Columns[14].ColumnWidth = 12;
                worksheet.Columns[15].ColumnWidth = 12;
                worksheet.Columns[16].ColumnWidth = 3.5;
                worksheet.Columns[17].ColumnWidth = 3.5;
                worksheet.Columns[18].ColumnWidth = 3.5;
                worksheet.Columns[19].ColumnWidth = 3.5;

                worksheet.Shapes.AddPicture(@"D:\EnvanterSistemi\efe-logo.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 26, 23, 45, 45);

                worksheet.Rows[1].RowHeight = 6;

                //((Range)worksheet.Range[(Range)worksheet.Cells[2, 2], (Range)worksheet.Cells[5, 18]]).Cells.Merge();
                ((Range)worksheet.Range[worksheet.Cells[2, 2], worksheet.Cells[5, 18]]).Cells.Merge();
                worksheet.Range["B2"].VerticalAlignment = XlVAlign.xlVAlignCenter;
                worksheet.Range["B2"].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                worksheet.Range["B2:R5"].Borders.Value = 1;
                worksheet.Cells[2, 2].Font.Size = 18;
                worksheet.Cells[2, 2] = "ZİMMET FORMU";

                worksheet.Rows[6].RowHeight = 6;

                for (int j = 7; j <= 9; j++)
                {
                    ((Range)worksheet.Range[(Range)worksheet.Cells[j, 2], (Range)worksheet.Cells[j, 3]]).Cells.Merge();
                    ((Range)worksheet.Range[(Range)worksheet.Cells[j, 2], (Range)worksheet.Cells[j, 3]]).Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    ((Range)worksheet.Range[(Range)worksheet.Cells[j, 2], (Range)worksheet.Cells[j, 3]]).Cells.HorizontalAlignment = XlHAlign.xlHAlignLeft;


                }
                for (int j = 7; j <= 9; j++)
                {
                    ((Range)worksheet.Range[(Range)worksheet.Cells[j, 4], (Range)worksheet.Cells[j, 18]]).Cells.Merge();
                    ((Range)worksheet.Range[(Range)worksheet.Cells[j, 4], (Range)worksheet.Cells[j, 18]]).Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    ((Range)worksheet.Range[(Range)worksheet.Cells[j, 4], (Range)worksheet.Cells[j, 18]]).Cells.HorizontalAlignment = XlHAlign.xlHAlignLeft;


                }

                for (int j = 10; j <= 15; j++) // this will apply it from col 1 to 10
                {
                    worksheet.Rows[j].RowHeight = 25;
                    ((Range)worksheet.Range[(Range)worksheet.Cells[j, 2], (Range)worksheet.Cells[j, 18]]).Cells.Merge();
                    ((Range)worksheet.Range[(Range)worksheet.Cells[j, 2], (Range)worksheet.Cells[j, 18]]).Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    ((Range)worksheet.Range[(Range)worksheet.Cells[j, 2], (Range)worksheet.Cells[j, 18]]).Cells.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                }
                worksheet.Range["B11"].VerticalAlignment = XlVAlign.xlVAlignCenter;
                worksheet.Range["B11"].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                worksheet.Range["B11:R13"].Borders.Value = 1;
                worksheet.Cells[11, 2] = "SORUMLULUKLAR";
                Excel.Range formatRange;
                formatRange = worksheet.get_Range("B11", "B11");
                formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);

                worksheet.Cells[7, 2] = "Adı Soyadı";
                worksheet.Cells[8, 2] = "Bölümü";
                worksheet.Cells[9, 2] = "Görevi";
                worksheet.Range["B7:C7"].Borders.Value = 1;
                worksheet.Range["B8:C8"].Borders.Value = 1;
                worksheet.Range["B9:C9"].Borders.Value = 1;
                worksheet.Range["D7:R7"].Borders.Value = 1;
                worksheet.Range["D8:R8"].Borders.Value = 1;
                worksheet.Range["D9:R9"].Borders.Value = 1;

                worksheet.Cells[7, 4] = adsoyad;
                worksheet.Cells[8, 4] = birim;
                worksheet.Cells[9, 4] = unvan;

                worksheet.Rows[10].RowHeight = 6;

                worksheet.Range["B12"].WrapText = true;
                worksheet.Cells[12, 2].Font.Size = 9;
                worksheet.Cells[12, 2] = "* Personel, iş akdinin herhangi bir nedenle sona ermesi halinde, zimmetindeki otomobil, bilgisayar, mobil cihazlar, telefon ve her türlü kayıtlı malzemeyi yöneticisine veya adı bildirilecek diğer personele teslim etmekle yükümlüdür.";

                worksheet.Range["B13"].WrapText = true;
                worksheet.Cells[13, 2].Font.Size = 9;
                worksheet.Cells[13, 2] = "* İş yerinde çalışana verilen donanımların içinde gelen tüm yazılımların işverenin yönetiminde olup, sadece iş için ve amaca yönelik olarak kullanılması gerekmektedir. Bu yazılımların ve mobil cihazların iş dışı amaçlarla kullanılması durumunda doğabilecek tüm hukuki süreçlerden çalışan sorumludur.";

                worksheet.Rows[14].RowHeight = 6;

                worksheet.Range["B15"].VerticalAlignment = XlVAlign.xlVAlignCenter;
                worksheet.Range["B15"].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                worksheet.Range["B15:R16"].Borders.Value = 1;
                worksheet.Cells[15, 2] = "EKİPMAN İLE İLGİLİ BİLGİLER";

                formatRange = worksheet.get_Range("B15", "R20");
                formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);

                //formatRange = worksheet.get_Range("B15", "B15");
                //formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                //formatRange = worksheet.get_Range("B16", "B16");
                //formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                worksheet.Rows[16].RowHeight = 25;
                ((Range)worksheet.Range[(Range)worksheet.Cells[16, 2], (Range)worksheet.Cells[16, 9]]).Cells.Merge();
                worksheet.Range["B16"].VerticalAlignment = XlVAlign.xlVAlignCenter;
                worksheet.Range["B16"].HorizontalAlignment = XlHAlign.xlHAlignLeft;

                worksheet.Range["B16:R31"].Borders.Value = 1;

                worksheet.Range["J16"].VerticalAlignment = XlVAlign.xlVAlignCenter;
                worksheet.Range["J16"].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                worksheet.Range["J16"].WrapText = true;
                worksheet.Cells[16, 10].Font.Size = 9;
                worksheet.Cells[16, 10] = "Teslim Alan (İmza)";

                worksheet.Range["K16"].VerticalAlignment = XlVAlign.xlVAlignCenter;
                worksheet.Range["K16"].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                worksheet.Range["K16"].WrapText = true;
                worksheet.Cells[16, 11].Font.Size = 9;
                worksheet.Cells[16, 11] = "Teslim Eden (İmza)";

                worksheet.Range["L16"].VerticalAlignment = XlVAlign.xlVAlignCenter;
                worksheet.Range["L16"].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                worksheet.Cells[16, 12] = "Onay";

                worksheet.Cells[16, 13] = "";

                worksheet.Range["N16"].VerticalAlignment = XlVAlign.xlVAlignCenter;
                worksheet.Range["N16"].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                worksheet.Range["N16"].WrapText = true;
                worksheet.Cells[16, 14].Font.Size = 9;
                worksheet.Cells[16, 14] = "İade Eden (İmza)";

                worksheet.Range["O16"].VerticalAlignment = XlVAlign.xlVAlignCenter;
                worksheet.Range["O16"].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                worksheet.Range["O16"].WrapText = true;
                worksheet.Cells[16, 15].Font.Size = 9;
                worksheet.Cells[16, 15] = "İade Alan (İmza)";

                worksheet.Range["P16"].VerticalAlignment = XlVAlign.xlVAlignCenter;
                worksheet.Range["P16"].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                ((Range)worksheet.Range[(Range)worksheet.Cells[16, 16], (Range)worksheet.Cells[16, 18]]).Cells.Merge();
                worksheet.Cells[16, 16] = "Onay";

                worksheet.Range["B17"].VerticalAlignment = XlVAlign.xlVAlignCenter;
                worksheet.Range["B17"].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                ((Range)worksheet.Range[(Range)worksheet.Cells[17, 2], (Range)worksheet.Cells[20, 5]]).Cells.Merge();
                worksheet.Cells[17, 2] = "Ürün Adı";

                worksheet.Range["F17"].VerticalAlignment = XlVAlign.xlVAlignCenter;
                worksheet.Range["F17"].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                ((Range)worksheet.Range[(Range)worksheet.Cells[17, 6], (Range)worksheet.Cells[20, 7]]).Cells.Merge();
                worksheet.Cells[17, 6] = "Marka Model";

                worksheet.Range["H17"].VerticalAlignment = XlVAlign.xlVAlignCenter;
                worksheet.Range["H17"].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                worksheet.Cells[17, 8] = "Seri No / IMEI";

                worksheet.Range["H18"].VerticalAlignment = XlVAlign.xlVAlignCenter;
                worksheet.Range["H18"].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                ((Range)worksheet.Range[(Range)worksheet.Cells[18, 8], (Range)worksheet.Cells[19, 8]]).Cells.Merge();
                worksheet.Cells[18, 8] = "Tel No / IMS No";

                worksheet.Range["H20"].VerticalAlignment = XlVAlign.xlVAlignCenter;
                worksheet.Range["H20"].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                worksheet.Cells[20, 8] = "Plaka No";

                worksheet.Range["I17"].VerticalAlignment = XlVAlign.xlVAlignCenter;
                worksheet.Range["I17"].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                ((Range)worksheet.Range[(Range)worksheet.Cells[17, 9], (Range)worksheet.Cells[20, 9]]).Cells.Merge();
                worksheet.Cells[17, 9] = "Teslim Tarihi";

                worksheet.Range["J17"].VerticalAlignment = XlVAlign.xlVAlignCenter;
                worksheet.Range["J17"].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                ((Range)worksheet.Range[(Range)worksheet.Cells[17, 10], (Range)worksheet.Cells[20, 10]]).Cells.Merge();
                worksheet.Cells[17, 10] = "Personel";




                worksheet.Range["K17"].VerticalAlignment = XlVAlign.xlVAlignCenter;
                worksheet.Range["K17"].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                ((Range)worksheet.Range[(Range)worksheet.Cells[17, 11], (Range)worksheet.Cells[18, 11]]).Cells.Merge();
                worksheet.Cells[17, 11] = "Bilgi İşlem";

                worksheet.Range["K19"].VerticalAlignment = XlVAlign.xlVAlignCenter;
                worksheet.Range["K19"].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                ((Range)worksheet.Range[(Range)worksheet.Cells[19, 11], (Range)worksheet.Cells[20, 11]]).Cells.Merge();
                worksheet.Cells[19, 11] = "İdari İşler";

                worksheet.Range["L17"].VerticalAlignment = XlVAlign.xlVAlignCenter;
                worksheet.Range["L17"].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                ((Range)worksheet.Range[(Range)worksheet.Cells[17, 12], (Range)worksheet.Cells[20, 12]]).Cells.Merge();
                worksheet.Cells[17, 12] = "Sıralı Amir";

                worksheet.Range["M17"].VerticalAlignment = XlVAlign.xlVAlignCenter;
                worksheet.Range["M17"].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                ((Range)worksheet.Range[(Range)worksheet.Cells[17, 13], (Range)worksheet.Cells[20, 13]]).Cells.Merge();
                worksheet.Cells[17, 13] = "İade Tarihi";

                worksheet.Range["N17"].VerticalAlignment = XlVAlign.xlVAlignCenter;
                worksheet.Range["N17"].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                ((Range)worksheet.Range[(Range)worksheet.Cells[17, 14], (Range)worksheet.Cells[20, 14]]).Cells.Merge();
                worksheet.Cells[17, 14] = "Personel";


                worksheet.Range["O17"].VerticalAlignment = XlVAlign.xlVAlignCenter;
                worksheet.Range["O17"].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                ((Range)worksheet.Range[(Range)worksheet.Cells[17, 15], (Range)worksheet.Cells[18, 15]]).Cells.Merge();
                worksheet.Cells[17, 15] = "Bilgi İşlem";

                worksheet.Range["O19"].VerticalAlignment = XlVAlign.xlVAlignCenter;
                worksheet.Range["O19"].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                ((Range)worksheet.Range[(Range)worksheet.Cells[19, 15], (Range)worksheet.Cells[20, 15]]).Cells.Merge();
                worksheet.Cells[19, 15] = "İdari İşler";

                worksheet.Range["P17"].VerticalAlignment = XlVAlign.xlVAlignCenter;
                worksheet.Range["P17"].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                ((Range)worksheet.Range[(Range)worksheet.Cells[17, 16], (Range)worksheet.Cells[20, 18]]).Cells.Merge();
                worksheet.Cells[17, 17] = "Sıralı Amir";


                worksheet.Range["B17"].VerticalAlignment = XlVAlign.xlVAlignCenter;
                worksheet.Range["B17"].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                worksheet.Range["B17:R20"].Borders.Value = 1;

                for (int j = 21; j <= 31; j++) // this will apply it from col 1 to 10
                {
                    worksheet.Rows[j].RowHeight = 28;
                    ((Range)worksheet.Range[(Range)worksheet.Cells[j, 2], (Range)worksheet.Cells[j, 5]]).Cells.Merge();
                    ((Range)worksheet.Range[(Range)worksheet.Cells[j, 2], (Range)worksheet.Cells[j, 5]]).Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    ((Range)worksheet.Range[(Range)worksheet.Cells[j, 2], (Range)worksheet.Cells[j, 5]]).Cells.HorizontalAlignment = XlHAlign.xlHAlignLeft;

                    ((Range)worksheet.Range[(Range)worksheet.Cells[j, 6], (Range)worksheet.Cells[j, 7]]).Cells.Merge();
                    ((Range)worksheet.Range[(Range)worksheet.Cells[j, 6], (Range)worksheet.Cells[j, 7]]).Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    ((Range)worksheet.Range[(Range)worksheet.Cells[j, 6], (Range)worksheet.Cells[j, 7]]).Cells.HorizontalAlignment = XlHAlign.xlHAlignLeft;

                    ((Range)worksheet.Range[(Range)worksheet.Cells[j, 16], (Range)worksheet.Cells[j, 18]]).Cells.Merge();
                    ((Range)worksheet.Range[(Range)worksheet.Cells[j, 16], (Range)worksheet.Cells[j, 18]]).Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    ((Range)worksheet.Range[(Range)worksheet.Cells[j, 16], (Range)worksheet.Cells[j, 18]]).Cells.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                }

                ((Range)worksheet.Range[(Range)worksheet.Cells[32, 2], (Range)worksheet.Cells[32, 5]]).Cells.Merge();
                ((Range)worksheet.Range[(Range)worksheet.Cells[32, 2], (Range)worksheet.Cells[32, 5]]).Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
                ((Range)worksheet.Range[(Range)worksheet.Cells[32, 2], (Range)worksheet.Cells[32, 5]]).Cells.HorizontalAlignment = XlHAlign.xlHAlignLeft;
                worksheet.Cells[32, 2] = "F-GN-005/R00/24.05.2023";

                ((Range)worksheet.Range[(Range)worksheet.Cells[32, 16], (Range)worksheet.Cells[32, 18]]).Cells.Merge();
                ((Range)worksheet.Range[(Range)worksheet.Cells[32, 16], (Range)worksheet.Cells[32, 18]]).Cells.VerticalAlignment = XlVAlign.xlVAlignCenter;
                ((Range)worksheet.Range[(Range)worksheet.Cells[32, 16], (Range)worksheet.Cells[32, 18]]).Cells.HorizontalAlignment = XlHAlign.xlHAlignRight;

                worksheet.Cells[32, 16].NumberFormat = "@";
                worksheet.Cells[32, 16] = "1/1";



                int rowstart = 21;
                int row = 21;
                int adet = ids.Length;
                int sonno = adet + rowstart;
                int i = 0;

                SatirDoldurma:
                foreach (string id in ids)
                {
                    if (row<sonno)
                    {

                        int zimmetid = Convert.ToInt32(ids[i]);

                        var zimmet = db.Zimmets.Where(x => x.id == zimmetid).FirstOrDefault();

                        var envanterid = zimmet.envanterId;
                        var envanterkodu = zimmet.envanterKodu;
                        var teslimtarihi = zimmet.teslimTarihi;
                        var digerperid = zimmet.personelId;


                        var envanterler = db.Envanters.Where(z => z.id == envanterid).FirstOrDefault();

                        int turid = Convert.ToInt32(envanterler.turId);
                        int kategoriid = Convert.ToInt32(envanterler.kategoriId);

                        var envanterdetay = db.EnvanterDetays.Where(y => y.envanterId == envanterid).FirstOrDefault();

                        var envanterturu = envanterdetay.envanterTuru;
                        int modelid = Convert.ToInt32(envanterdetay.modelId);

                        var markamodel = db.MarkaModels.Where(x => x.id == modelid).FirstOrDefault();
                        var marka = markamodel.marka;
                        var model = markamodel.model;
                        var serino = envanterdetay.seriNo;

                        string mm = marka + "/" + model;
                        string digerbilgi = "";

                        //  ,[envanterTuru]
                        //  ,[modelId]
                        //  ,[seriNo]

                        if (turid==19 || turid == 21 || turid == 22 || turid == 23 || turid == 25 || turid == 133 || turid == 134 || turid == 143)
                        {//bilgisayar
                            var pcturu = envanterdetay.pcTuru;
                            var islemci = envanterdetay.islemci;
                            var hdd = envanterdetay.hdd;
                            var ssd = envanterdetay.ssd;
                            var ram = envanterdetay.ram;
                            var ekrankarti = envanterdetay.ek;

                            var lanmac = envanterdetay.lanMac;
                            var wifimac = envanterdetay.wifiMac;
                            var imei1 = envanterdetay.imei1;
                            var imei2 = envanterdetay.imei2;
                            //  ,[pcTuru]
                            //  ,[islemci]
                            //  ,[hdd]
                            //  ,[ssd]
                            //  ,[ram]
                            //  ,[ek]
                            digerbilgi =serino;

                        }
                        else if (turid == 125)
                        {//personel kartı
                            //  ,[onyuzNo]
                            //  ,[arkayuzNo]
                            //  ,[kgbNo]
                            var onyuzno = envanterdetay.onyuzNo;
                            var arkayuzno = envanterdetay.arkayuzNo;
                            var kgbno = envanterdetay.kgbNo;
                            digerbilgi = onyuzno + "/" + arkayuzno+"/"+kgbno;
                        }
                        else if (turid == 83 || turid == 84 || turid == 85 || turid == 86 || turid == 87 || turid == 88 || turid == 89 || turid == 138 || turid == 139)
                        {//mobilya
                            var mobilyacinsi = envanterdetay.mobilyaCinsi;
                            var cekmece = envanterdetay.cekmece;
                            var kapak = envanterdetay.kapak;
                            var en = envanterdetay.en;
                            var boy = envanterdetay.boy;
                            var yukseklik = envanterdetay.yukseklik;
                            //  ,[mobilyaCinsi]
                            //  ,[cekmece]
                            //  ,[kapak]
                            //  ,[en]
                            //  ,[boy]
                            //  ,[yukseklik]
                            digerbilgi = envanterturu;
                        }
                        else if (turid == 1 || turid == 2 || turid == 3 || turid == 4)
                        {//araba
                            //  ,[plaka]
                            var plaka = envanterdetay.plaka;
                            digerbilgi = plaka;
                        }
                        else if (turid == 10 || turid == 11 || turid == 12 || turid == 13 || turid == 14 || turid == 15 || turid == 16 || turid == 17 || turid == 18 || turid == 74 || turid == 135)
                        {//hatlar
                         //   ,[hatNo]
                         //  ,[simId]
                           
                            var hatno = envanterdetay.hatNo;
                            var simid = envanterdetay.simId;
                            var lanmac = envanterdetay.lanMac;
                            var wifimac = envanterdetay.wifiMac;
                            var imei1 = envanterdetay.imei1;
                            var imei2 = envanterdetay.imei2;
                            digerbilgi =hatno+"/"+simid;
                        }




                        worksheet.Cells[row, 2].VerticalAlignment = XlVAlign.xlVAlignCenter;
                        worksheet.Cells[row, 2].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                        worksheet.Cells[row, 2].WrapText = true;
                        worksheet.Cells[row, 2].Font.Size = 9;
                        worksheet.Cells[row, 2] = envanterkodu + "/" + envanterturu;
                        worksheet.Cells[row, 6].VerticalAlignment = XlVAlign.xlVAlignCenter;
                        worksheet.Cells[row, 6].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                        worksheet.Cells[row, 6].WrapText = true;
                        worksheet.Cells[row, 6].Font.Size = 9;
                        worksheet.Cells[row, 6] = marka + "/" + model;
                        worksheet.Cells[row, 8].VerticalAlignment = XlVAlign.xlVAlignCenter;
                        worksheet.Cells[row, 8].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                        worksheet.Cells[row, 8].WrapText = true;
                        worksheet.Cells[row, 8].Font.Size = 9;
                        worksheet.Cells[row, 8] = digerbilgi;
                        worksheet.Cells[row, 9].VerticalAlignment = XlVAlign.xlVAlignCenter;
                        worksheet.Cells[row, 9].HorizontalAlignment = XlHAlign.xlHAlignLeft;
                        worksheet.Cells[row, 9].WrapText = true;
                        worksheet.Cells[row, 9].Font.Size = 9;
                        worksheet.Cells[row, 9] = teslimtarihi;
                        Range rg = (Excel.Range)worksheet.Cells[row, 9];
                        rg.EntireColumn.NumberFormat = "dd.mm.yyyy";



                        row++;
                        i++;
                        goto SatirDoldurma;
                    }
                    else
                    {
                        TempData["Mesaj"] = "İşlem Tamamlandı";
                    }

                    
                }

                string kayitzamani2 = DateTime.Now.ToString("yyyyMMddHHmmss");
                workbook.SaveAs("D:\\EnvanterSistemi\\ZimmetFormu" + username + "-" + kayitzamani2 + ".xlsx");
                #endregion

                TempData["Mesaj"] = "D:\\EnvanterSistemi\\ZimmetFormu" + username + "-" + kayitzamani2 + ".xlsx" + "  zimmet formu oluşturulmuştur.";

                workbook.Close();
                application.Quit();
            }
            
            return RedirectToAction("BaskiAl","Zimmet");
        }


        [HttpPost]
        public JsonResult PersonelGetir(int? p)
        {

            int birimid = Convert.ToInt32(p);
            var turler = (from x in db.Personels
                          join y in db.Birimlers on x.birimAdi equals y.birimAdi
                          where x.konumId==birimid
                          select new
                          {
                              Text = x.adSoyad,
                              Value = x.id.ToString()
                          }).ToList();

            return Json(turler, JsonRequestBehavior.AllowGet);
        }
        [HttpPost]
        public JsonResult OdaGetir(string p)
        {

            string lokasyon = Convert.ToString(p);
            var turler = (from x in db.Konums
                          where x.lokasyon == lokasyon
                          select new
                          {
                              Text = x.odaNo,
                              Value = x.odaNo
                          }).ToList();

            return Json(turler, JsonRequestBehavior.AllowGet);
        }
    }
}