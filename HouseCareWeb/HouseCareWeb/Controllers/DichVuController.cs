using HouseCareWeb.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace HouseCareWeb.Controllers
{
    public class DichVuController : Controller
    {
        // GET: DichVu
        public ActionResult Index()
        {
            return View();
        }
        public ActionResult Viecnha()
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
            List<goidichvu> goidichvuList = dbcontext.goidichvus.Where(row => row.Madichvu == 1).ToList();
            ViewBag.goidichvu = goidichvuList;
            return View();
        }
        public ActionResult Vesinhcongnghiep()
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
            List<goidichvu> goidichvuVesinhcongnghiep = dbcontext.goidichvus.Where(row => row.Madichvu == 2).ToList();
            ViewBag.goidichvu = goidichvuVesinhcongnghiep;
            return View();
        }
        public ActionResult Suachuadienlanh()
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
            List<goidichvu> goidichvusuachuadienlanh = dbcontext.goidichvus.Where(row => row.Madichvu == 3).ToList();
            ViewBag.goidichvu = goidichvusuachuadienlanh;
            return View();
        }
        public ActionResult Chamsocmevabe()
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
            List<goidichvu> goidichvucsmevabe = dbcontext.goidichvus.Where(row => row.Madichvu == 4).ToList();
            ViewBag.goidichvu = goidichvucsmevabe;
            return View();
        }
    }
}