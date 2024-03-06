using HouseCareWeb.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace HouseCareWeb.Controllers
{
    public class CheckoutController : Controller
    {
        // GET: Checkout
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult CheckoutLogin()
        {
            housecareDBContextDataContext context = new housecareDBContextDataContext();
            int idgoi = -1;
            try
            {
                idgoi = int.Parse(Request.QueryString["id"]);
            }
            catch
            {
                idgoi = -1;
            }
            ViewData["idmagoi"] = idgoi;
            account tk = Session["Taikhoan"] as HouseCareWeb.Models.account;

            goidichvu goidichvu = context.goidichvus.FirstOrDefault(x => x.Id == idgoi);
            ViewBag.goidichvudachon = goidichvu;
            List<dichvu> dv = context.dichvus.ToList();
            ViewBag.dv = dv;

            return View();
        }

        public ActionResult CheckoutNoLogin()
        {
            housecareDBContextDataContext context = new housecareDBContextDataContext();
            int idgoi = -1;
            try
            {
                idgoi = int.Parse(Request.QueryString["id"]);
            }
            catch
            {
                idgoi = -1;
            }
            ViewData["idmagoi"] = idgoi;

            goidichvu goidichvu = context.goidichvus.FirstOrDefault(x => x.Id == idgoi);
            ViewBag.goidichvudachon = goidichvu;
            List<dichvu> dv = context.dichvus.ToList();
            ViewBag.dv = dv;
            return View();
        }

        [HttpPost]
        public ActionResult ConfirmCheckoutNoLogin()
        {
            int idgoi = -1;
            try
            {
                idgoi = int.Parse(Request.QueryString["id"]);
            }
            catch
            {
                idgoi = -1;
            }
            ViewData["idmagoi"] = idgoi;

            if (Request.Form.Count > 0)
            {
                housecareDBContextDataContext context = new housecareDBContextDataContext();
                int idmagoi = int.Parse(Request.Form["idgoi"]);
                String ten = Request.Form["Hoten"];
                String sdt = Request.Form["Sdt"];
                String email = Request.Form["Email"];
                String diachi = Request.Form["Diachi"];
                String yeucau = Request.Form["Yeucau"];
                DateTime ngaydangky = DateTime.Parse(Request.Form["Ngaydangky"]);
                String gio = Request.Form["Gio"];

                account tk = new account();

                tk = context.accounts.FirstOrDefault(x => x.Sdt == sdt);
                goidichvu goidichvu = context.goidichvus.FirstOrDefault(x => x.Id == idmagoi);

                if (tk == null)
                {
                    account tktemp = new account();
                    tktemp.Ten = ten;
                    tktemp.Sdt = sdt;
                    tktemp.Email = email;
                    tktemp.Taikhoan = null;
                    tktemp.Matkhau = null;
                    tktemp.Quyen = "ND";
                    context.accounts.InsertOnSubmit(tktemp);
                    context.SubmitChanges();
                }
                tk = context.accounts.FirstOrDefault(x => x.Sdt == sdt);
                dangkygoidichvu dangkygoidichvu = new dangkygoidichvu();
                dangkygoidichvu.Magoi = goidichvu.Id;
                dangkygoidichvu.Idtaikhoan = tk.Id;
                dangkygoidichvu.Ngaydangky = ngaydangky;
                dangkygoidichvu.Diachi = diachi;
                dangkygoidichvu.Yeucau = yeucau;
                dangkygoidichvu.Trangthai = false;
                dangkygoidichvu.Active = false;
                dangkygoidichvu.Thoiluong = 1;

                context.dangkygoidichvus.InsertOnSubmit(dangkygoidichvu);
                context.SubmitChanges();
                return RedirectToAction("CheckoutSuccess");
            }
            return View();
        }
        [HttpPost]
        public ActionResult ConfirmCheckoutLogin()
        {
            housecareDBContextDataContext context = new housecareDBContextDataContext();
            int idgoi = -1;
            try
            {
                idgoi = int.Parse(Request.QueryString["id"]);
            }
            catch
            {
                idgoi = -1;
            }
            ViewData["idmagoi"] = idgoi;

            if (Request.Form.Count > 0)
            {
                DateTime ngaydangky = DateTime.Parse(Request.Form["Ngaydangky"]);
                String diachi = Request.Form["Diachi"];
                String yeucau = Request.Form["Yeucau"];

                account tk = Session["Taikhoan"] as HouseCareWeb.Models.account;
                int idmagoi = int.Parse(Request.Form["idgoi"]);

                goidichvu goidichvu = context.goidichvus.FirstOrDefault(x => x.Id == idmagoi);

                dangkygoidichvu dangkygoidichvu = new dangkygoidichvu();

                dangkygoidichvu.Magoi = goidichvu.Id;
                dangkygoidichvu.Idtaikhoan = tk.Id;
                dangkygoidichvu.Ngaydangky = ngaydangky;
                dangkygoidichvu.Diachi = diachi;
                dangkygoidichvu.Yeucau = yeucau;
                dangkygoidichvu.Trangthai = false;
                dangkygoidichvu.Active = false;
                dangkygoidichvu.Thoiluong = 1;
                context.dangkygoidichvus.InsertOnSubmit(dangkygoidichvu);
                context.SubmitChanges();
                return RedirectToAction("CheckoutSuccess");
            }
            return View();
        }

        public ActionResult CheckoutSuccess()
        {
            return View();
        }

    }

}