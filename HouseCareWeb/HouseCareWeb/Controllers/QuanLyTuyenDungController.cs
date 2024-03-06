using HouseCareWeb.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace HouseCareWeb.Controllers
{
    [Authorize(Users = "admin, nhanvien")]

    public class QuanLyTuyenDungController : Controller
    {
        // GET: QuanLyTuyenDung
        public ActionResult Index(String search = "", String sortColumn = "Id", String iconClass = "fa-solid fa-sort-up", int page = 1)
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
            List<Tuyendung> acc = dbcontext.Tuyendungs.Where(row => row.Hoten.Contains(search)).ToList();
            ViewBag.search = search;

            // Sort

            ViewBag.sortColumn = sortColumn;
            ViewBag.iconClass = iconClass;

            if (sortColumn == "Id")
            {
                if (iconClass == "fa-solid fa-sort-up")
                {
                    acc = acc.OrderBy(row => row.Id).ToList();
                }
                else
                {
                    acc = acc.OrderByDescending(row => row.Id).ToList();
                }
            }
            else if (sortColumn == "Hoten")
            {
                if (iconClass == "fa-solid fa-sort-up")
                {
                    acc = acc.OrderBy(row => row.Hoten).ToList();
                }
                else
                {
                    acc = acc.OrderByDescending(row => row.Hoten).ToList();
                }
            }
            else if (sortColumn == "Diachi")
            {
                if (iconClass == "fa-solid fa-sort-up")
                {
                    acc = acc.OrderBy(row => row.Diachi).ToList();
                }
                else
                {
                    acc = acc.OrderByDescending(row => row.Diachi).ToList();
                }
            }

            // Page

            int NoOfRecordPerPage = 5;
            int NoOfPages = Convert.ToInt32(Math.Ceiling(Convert.ToDouble(acc.Count) / Convert.ToDouble(NoOfRecordPerPage)));
            int NoOfRecordToSkip = (page - 1) * NoOfRecordPerPage;

            ViewBag.Page = page;
            ViewBag.NoOfPages = NoOfPages;

            acc = acc.Skip(NoOfRecordToSkip).Take(NoOfRecordPerPage).ToList();
            return View(acc);
        }
        public ActionResult Duyet(int id)
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
            Tuyendung tuyendung = dbcontext.Tuyendungs.FirstOrDefault(row => row.Id == id);

            if (tuyendung != null)
            {
                tuyendung.Trangthai = true;

                dbcontext.SubmitChanges();
            }
            return RedirectToAction("Index");
        }
    }
   
}