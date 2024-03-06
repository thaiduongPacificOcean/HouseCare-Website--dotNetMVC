using HouseCareWeb.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace HouseCareWeb.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
            List<dichvu> dichvuList = dbcontext.dichvus.ToList();
            
            for (int i = 0; i < dichvuList.Count; i++)
            {
                if (i > 5)
                {
                    dichvuList.RemoveAt(i);
                }
            }
            ViewBag.dichvu = dichvuList;

            List<Post> Listbaiviet = dbcontext.Posts.ToList();

            //for (int i = 0; i < Listbaiviet.Count; i++)
            //{
            //    if (i > 3)
            //    { 
            //        Listbaiviet.RemoveAt(i);
            //    }
            //}
            ViewBag.baiviet = Listbaiviet;

            return View();
        }

        public ActionResult About()
        {
            
            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
        public ActionResult Tuyendung()
        {
            return View();
        }
    }
}