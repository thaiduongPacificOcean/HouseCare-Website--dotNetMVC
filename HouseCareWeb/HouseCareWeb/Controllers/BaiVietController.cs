using HouseCareWeb.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Services.Description;

namespace HouseCareWeb.Controllers
{
    [Authorize(Users = "admin, nhanvien")]

    public class BaiVietController : Controller
    {
        // GET: BaiViet
        public ActionResult Index(String search = "", String sortColumn = "Id", String iconClass = "fa-solid fa-sort-up", int page = 1)
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();

            // Find

            List<Post> postList = dbcontext.Posts.Where(row => row.Ten.Contains(search)).ToList();
            ViewBag.search = search;

            // Sort

            ViewBag.sortColumn = sortColumn;
            ViewBag.iconClass = iconClass;

            if (sortColumn == "Id")
            {
                if (iconClass == "fa-solid fa-sort-up")
                {
                    postList = postList.OrderBy(row => row.Id).ToList();
                }
                else
                {
                    postList = postList.OrderByDescending(row => row.Id).ToList();
                }
            }
            
            else if (sortColumn == "Ten")
            {
                if (iconClass == "fa-solid fa-sort-up")
                {
                    postList = postList.OrderBy(row => row.Ten).ToList();
                }
                else
                {
                    postList = postList.OrderByDescending(row => row.Ten).ToList();
                }
            }

            // Page

            int NoOfRecordPerPage = 5;
            int NoOfPages = Convert.ToInt32(Math.Ceiling(Convert.ToDouble(postList.Count) / Convert.ToDouble(NoOfRecordPerPage)));
            int NoOfRecordToSkip = (page - 1) * NoOfRecordPerPage;

            ViewBag.Page = page;
            ViewBag.NoOfPages = NoOfPages;

            postList = postList.Skip(NoOfRecordToSkip).Take(NoOfRecordPerPage).ToList();

            return View(postList);
        }
        public ActionResult Edit(int id)
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
            Post post = dbcontext.Posts.FirstOrDefault(row => row.Id == id);

            string preimg = post.Img;


            if (Request.Form.Count > 0)
            {
                String ten = Request.Form["Ten"];
                String ngaydang = Request.Form["Ngaydang"];
                String mota = Request.Form["Mota"];

                post.Ten = ten;
                post.Mota = mota;
                post.Ngaydang = ngaydang;

                HttpPostedFileBase file = Request.Files["Img"];

                if (file != null)
                {
                    string serverpath = HttpContext.Server.MapPath("~/img");
                    string filepath = serverpath + "/" + file.FileName;
                    if (file.FileName.Equals(""))
                    {
                        post.Img = preimg;
                    }
                    else
                    {
                        file.SaveAs(filepath);
                        post.Img = file.FileName;
                    }
                }
                dbcontext.SubmitChanges();
                return RedirectToAction("Index");
            }
            return View(post);
        }

        public ActionResult Create()
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
            if (Request.Form.Count > 0)
            {
                String ten = Request.Form["Ten"];
                String ngaydang = Request.Form["Ngaydang"];
                String mota = Request.Form["Mota"];

                Post post = new Post();
                post.Ten = ten;
                post.Mota = mota;
                post.Ngaydang = ngaydang;

                HttpPostedFileBase file = Request.Files["Img"];

                if (file != null)
                {
                    string serverpath = HttpContext.Server.MapPath("~/img");
                    string filepath = serverpath + "/" + file.FileName;
                    if (file.FileName.Equals(""))
                    {
                        post.Img = "";
                    }
                    else
                    {
                        file.SaveAs(filepath);
                        post.Img = file.FileName;
                    }
                }
                dbcontext.Posts.InsertOnSubmit(post);
                dbcontext.SubmitChanges();
                return RedirectToAction("Index");
            }
            return View();
        }
        public ActionResult Delete(int id)
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
            Post post = dbcontext.Posts.FirstOrDefault(row => row.Id == id);
            return View(post);
        }
        [HttpPost]
        public ActionResult Delete(int id, nhanvien nv)
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
            Post post = dbcontext.Posts.FirstOrDefault(row => row.Id == id);
            if (post != null)
            {
                dbcontext.Posts.DeleteOnSubmit(post);
                dbcontext.SubmitChanges();
                return RedirectToAction("Index");
            }
            return View();
        }


        public ActionResult CSMVB()
        {
            return View();
        }
    }
}