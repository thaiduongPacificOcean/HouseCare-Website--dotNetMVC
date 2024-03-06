using ClosedXML.Excel;
using HouseCareWeb.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace HouseCareWeb.Controllers
{
    [Authorize(Users = "admin")]
    public class QuanLyTaiKhoanController : Controller
    {
        // GET: QuanLyTaiKhoan
        public ActionResult Index(String search="", String sortColumn = "Id", String iconClass = "fa-solid fa-sort-up", int page = 1)
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
            List<account> acc = dbcontext.accounts.Where(row => row.Taikhoan.Contains(search)).ToList();
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
            else if (sortColumn == "Mataikhoan")
            {
                if (iconClass == "fa-solid fa-sort-up")
                {
                    acc = acc.OrderBy(row => row.Mataikhoan).ToList();
                }
                else
                {
                    acc = acc.OrderByDescending(row => row.Mataikhoan).ToList();
                }
            }
            else if (sortColumn == "Taikhoan")
            {
                if (iconClass == "fa-solid fa-sort-up")
                {
                    acc = acc.OrderBy(row => row.Taikhoan).ToList();
                }
                else
                {
                    acc = acc.OrderByDescending(row => row.Taikhoan).ToList();
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
        public ActionResult Details(int id)
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
            account acc = dbcontext.accounts.FirstOrDefault(row => row.Id == id);
            return View(acc);
        }
        public ActionResult Edit(int id)
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
            account acc = dbcontext.accounts.FirstOrDefault(row => row.Id == id);
            string preimg = acc.Img;    

            if (Request.Form.Count > 0)
            {
                String taikhoan = Request.Form["Taikhoan"];
                String email = Request.Form["Email"];
                String matkhau = Request.Form["Matkhau"];
                String ten = Request.Form["Ten"];
                String sdt = Request.Form["Sdt"];
                String quyen = Request.Form["Quyen"];

                acc.Taikhoan = taikhoan;
                acc.Email = email;
                acc.Matkhau = matkhau;
                acc.Ten = ten;
                acc.Sdt = sdt;
                acc.Quyen = quyen;

                HttpPostedFileBase file = Request.Files["Img"];

                if (file != null)
                {
                    string serverpath = HttpContext.Server.MapPath("~/img");
                    string filepath = serverpath + "/" + file.FileName;
                    if (file.FileName.Equals(""))
                    {
                        acc.Img = preimg;
                    }
                    else
                    {
                        file.SaveAs(filepath);
                        acc.Img = file.FileName;
                    }
                }
                dbcontext.SubmitChanges();
                return RedirectToAction("Index");
            }
            return View(acc);
        }

        public ActionResult Create()
        {
            if (Request.Form.Count > 0)
            {
                String matk = Request.Form["Mataikhoan"];
                String taikhoan = Request.Form["Taikhoan"];
                String email = Request.Form["Email"];
                String matkhau = Request.Form["Matkhau"];
                String ten = Request.Form["Ten"];
                String sdt = Request.Form["Sdt"];
                String quyen = Request.Form["Quyen"];
                housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
                account acc = new account();
                acc.Mataikhoan = "user"+matk;
                acc.Taikhoan = taikhoan;
                acc.Email = email;
                acc.Matkhau = matkhau;
                acc.Ten = ten;
                acc.Sdt = sdt;
                acc.Quyen= "ND";
                HttpPostedFileBase file = Request.Files["Img"];

                if (file != null)
                {
                    string serverpath = HttpContext.Server.MapPath("~/img");
                    string filepath = serverpath + "/" + file.FileName;
                    if (file.FileName.Equals(""))
                    {
                        acc.Img = "";
                    }
                    else
                    {
                        file.SaveAs(filepath);
                        acc.Img = file.FileName;
                    }
                }
                dbcontext.accounts.InsertOnSubmit(acc);
                dbcontext.SubmitChanges();
                return RedirectToAction("Index");
            }
            return View();
        }
        public ActionResult Delete(int id)
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
            account acc = dbcontext.accounts.FirstOrDefault(row => row.Id == id);
            return View(acc);   
        }
        [HttpPost]
        public ActionResult Delete(int id ,account tk)
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
            account acc = dbcontext.accounts.FirstOrDefault(row => row.Id == id);
            if (acc != null)
            {
                dbcontext.accounts.DeleteOnSubmit(acc);
                dbcontext.SubmitChanges();
            }
            return RedirectToAction("Index");
        }
        // Export Excel 
        public ActionResult ExportToExcel()
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();

            List<account> all = dbcontext.accounts.ToList();

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Students");
                var currentRow = 1;
                worksheet.Cell(currentRow, 1).Value = "Id";
                worksheet.Cell(currentRow, 2).Value = "Mã tài khoản";
                worksheet.Cell(currentRow, 3).Value = "Tài khoản";
                worksheet.Cell(currentRow, 4).Value = "Email";
                worksheet.Cell(currentRow, 5).Value = "Mật khẩu";
                worksheet.Cell(currentRow, 6).Value = "Tên";
                worksheet.Cell(currentRow, 7).Value = "Số điện thoại";
                worksheet.Cell(currentRow, 8).Value = "Quyền";
                worksheet.Cell(currentRow, 9).Value = "Hình ảnh";

                foreach (var nv in all)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = nv.Id;
                    worksheet.Cell(currentRow, 2).Value = nv.Mataikhoan;
                    worksheet.Cell(currentRow, 3).Value = nv.Taikhoan;
                    worksheet.Cell(currentRow, 4).Value = nv.Email;
                    worksheet.Cell(currentRow, 5).Value = nv.Matkhau;
                    worksheet.Cell(currentRow, 6).Value = nv.Ten;
                    worksheet.Cell(currentRow, 7).Value = nv.Sdt;
                    worksheet.Cell(currentRow, 8).Value = nv.Quyen;
                    worksheet.Cell(currentRow, 9).Value = nv.Img;


                }
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();
                    return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Danhsachtaikhoan.xlsx");
                }
            }
        }
    }
}