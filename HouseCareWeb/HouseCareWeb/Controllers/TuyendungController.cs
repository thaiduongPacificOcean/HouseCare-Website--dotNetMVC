using HouseCareWeb.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace HouseCareWeb.Controllers
{
    public class TuyendungController : Controller
    {
        // GET: Tuyendung
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult Thongtintuyendung()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Dangkytuyendung()
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
            if (Request.Form.Count > 0)
            {
                String ten = Request.Form["Hoten"];
                String gioitinh = Request.Form["Gioitinh"];
                String diachi = Request.Form["Diachi"];
                String quequan = Request.Form["Quequan"];
                String chucvu = Request.Form["ChucVu"];
                //int tuoi = int.Parse(Request.Form["Tuoi"]);

                DateTime ngaysinh = DateTime.Parse(Request.Form["Ngaysinh"]);
                DateTime ngaydangky = DateTime.Parse(Request.Form["Ngaydangky"]);


                String cccd = Request.Form["CCCD"];
                String email = Request.Form["Email"];
                String sdt = Request.Form["Sdt"];
                String mota = Request.Form["Mota"];
                String kinhnghiem = Request.Form["Kinhnghiem"];

                Tuyendung tuyendung = new Tuyendung();
                

                tuyendung.Hoten = ten;
                tuyendung.Gioitinh = gioitinh;
                tuyendung.Ngaysinh = ngaysinh;
                tuyendung.Diachi = diachi;
                tuyendung.Quequan = quequan;
                tuyendung.Chucvu = chucvu;
                tuyendung.CCCD = cccd;
                tuyendung.Email = email;
                tuyendung.Sdt = sdt;
                tuyendung.Mota = mota;
                tuyendung.Kinhnghiem = kinhnghiem;
                tuyendung.Ngaydangky = DateTime.Now;
                tuyendung.Trangthai = false;
                HttpPostedFileBase file = Request.Files["Img"];

                if (file != null)
                {
                    string serverpath = HttpContext.Server.MapPath("~/img");
                    string filepath = serverpath + "/" + file.FileName;
                    if (file.FileName.Equals(""))
                    {
                        tuyendung.Img = "";
                    }
                    else
                    {
                        file.SaveAs(filepath);
                        tuyendung.Img = file.FileName;
                    }
                }
                dbcontext.Tuyendungs.InsertOnSubmit(tuyendung);
                dbcontext.SubmitChanges();
                return RedirectToAction("Thongbaotuyendung");
            }
            return View();
        }
        public ActionResult Thongbaotuyendung()
        {
            return View();
        }
        
    }
}