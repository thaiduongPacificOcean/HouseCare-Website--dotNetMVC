using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace HouseCareWeb.Controllers
{
    [Authorize(Users = "admin, nhanvien")]
    public class QuanLyChiTietDichVuController : Controller
    {
        // GET: QuanLyChiTietDichVu
        public ActionResult Index()
        {
            return View();
        }
    }
}