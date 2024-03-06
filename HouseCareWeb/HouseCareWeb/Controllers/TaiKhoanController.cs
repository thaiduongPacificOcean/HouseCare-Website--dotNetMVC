using HouseCareWeb.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace HouseCareWeb.Controllers
{
    public class TaiKhoanController : Controller
    {
        // GET: TaiKhoan

        public ActionResult Index()
        {
            housecareDBContextDataContext context = new housecareDBContextDataContext();

            account tk = Session["Taikhoan"] as HouseCareWeb.Models.account;

            List<dangkygoidichvu> dkgdv = context.dangkygoidichvus.Where(x => x.Idtaikhoan == tk.Id).ToList(); 
            
            ViewBag.listdangkygoidichvu = dkgdv;
            
            ViewBag.listgoidicvu = context.goidichvus.ToList();

            ViewBag.loaidichvu = context.dichvus.ToList();
            //
            

            return View();
        }
        public ActionResult ChangePass(int id)
        {

            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();

            return View();
        }
        [Authorize(Users = "admin, nhanvien")]
        public ActionResult AdminAccount( String sortColumn = "Id", String iconClass = "fa-solid fa-sort-up", int page = 1)
        {
            housecareDBContextDataContext context = new housecareDBContextDataContext();

            List<dangkygoidichvu> dkgdv = context.dangkygoidichvus.Where(x => x.Trangthai == false).ToList();


            

            // Sort

            int NoOfRecordPerPage = 5;
            int NoOfPages = Convert.ToInt32(Math.Ceiling(Convert.ToDouble(dkgdv.Count) / Convert.ToDouble(NoOfRecordPerPage)));
            int NoOfRecordToSkip = (page - 1) * NoOfRecordPerPage;

            ViewBag.Page = page;
            ViewBag.NoOfPages = NoOfPages;

            dkgdv = dkgdv.Skip(NoOfRecordToSkip).Take(NoOfRecordPerPage).ToList();
            ViewBag.listdangkygoidichvu = dkgdv;
            ViewBag.listgoidicvu = context.goidichvus.ToList();
            ViewBag.loaidichvu = context.dichvus.ToList();

            return View(dkgdv);
        }
        public ActionResult Duyet(int id)
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();

            dangkygoidichvu dangkygoidichvu = dbcontext.dangkygoidichvus.FirstOrDefault(row => row.Id == id);

            if (dangkygoidichvu != null)
            {
                dangkygoidichvu.Trangthai = true;

                dbcontext.SubmitChanges();
            }
            return View("AdminAccount");
        }
        
    }
}