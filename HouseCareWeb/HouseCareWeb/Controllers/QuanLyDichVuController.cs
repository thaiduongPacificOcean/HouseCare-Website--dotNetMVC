using ClosedXML.Excel;
using HouseCareWeb.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;


namespace HouseCareWeb.Controllers
{
    [Authorize(Users = "admin, nhanvien")]
    public class QuanLyDichVuController : Controller
    {
        // GET: QuanLyDichVu
        public ActionResult Index(string search="", String sortColumn = "Id", String iconClass = "fa-solid fa-sort-up", int page = 1)
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
            List<dichvu> dichvu = dbcontext.dichvus.Where(row => row.Tendichvu.Contains(search)).ToList();
            ViewBag.search = search;
            // Sort

            ViewBag.sortColumn = sortColumn;
            ViewBag.iconClass = iconClass;

            if (sortColumn == "Id")
            {
                if (iconClass == "fa-solid fa-sort-up")
                {
                    dichvu = dichvu.OrderBy(row => row.Id).ToList();
                }
                else
                {
                    dichvu = dichvu.OrderByDescending(row => row.Id).ToList();
                }
            }
            else if (sortColumn == "Madichvu")
            {
                if (iconClass == "fa-solid fa-sort-up")
                {
                    dichvu = dichvu.OrderBy(row => row.Madichvu).ToList();
                }
                else
                {
                    dichvu = dichvu.OrderByDescending(row => row.Madichvu).ToList();
                }
            }
            else if (sortColumn == "Tendichvu")
            {
                if (iconClass == "fa-solid fa-sort-up")
                {
                    dichvu = dichvu.OrderBy(row => row.Tendichvu).ToList();
                }
                else
                {
                    dichvu = dichvu.OrderByDescending(row => row.Tendichvu).ToList();
                }
            }
            
            // Page

            int NoOfRecordPerPage = 5;
            int NoOfPages = Convert.ToInt32(Math.Ceiling(Convert.ToDouble(dichvu.Count) / Convert.ToDouble(NoOfRecordPerPage)));
            int NoOfRecordToSkip = (page - 1) * NoOfRecordPerPage;

            ViewBag.Page = page;
            ViewBag.NoOfPages = NoOfPages;

            dichvu = dichvu.Skip(NoOfRecordToSkip).Take(NoOfRecordPerPage).ToList();
            return View(dichvu);
        }
        public ActionResult Details(int id)
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
            dichvu dichvu = dbcontext.dichvus.FirstOrDefault(row => row.Id == id);
            return View(dichvu);
        }
        public ActionResult Edit(int id)
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
            dichvu dichvu = dbcontext.dichvus.FirstOrDefault(row => row.Id == id);
            string preimg = dichvu.Img;

            if (Request.Form.Count > 0)
            {

                String tendichvu = Request.Form["Tendichvu"];
                String mota = Request.Form["Mota"];
                HttpPostedFileBase file = Request.Files["Img"];

                dichvu.Tendichvu = tendichvu;
                dichvu.Mota = mota;
                if (file != null)
                {
                    string serverpath = HttpContext.Server.MapPath("~/img");
                    string filepath = serverpath + "/" +file.FileName;
                    if (file.FileName.Equals(""))
                    {
                        dichvu.Img = preimg;
                    }
                    else
                    {
                        file.SaveAs(filepath);
                        dichvu.Img = file.FileName;
                    }

                }
                dbcontext.SubmitChanges();
                return RedirectToAction("Index");
            }
            return View(dichvu);
        }
        
        public ActionResult Create()
        {
            if (Request.Form.Count > 0)
            {
                String madichvu = Request.Form["Madichvu"];
                String tendichvu = Request.Form["Tendichvu"];
                String mota = Request.Form["Mota"];

                housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();

                dichvu dichvu = new dichvu();

                dichvu.Madichvu = "DV" + madichvu;
                dichvu.Tendichvu = tendichvu;
                dichvu.Mota = mota;

                HttpPostedFileBase file = Request.Files["Img"];
                if (file != null)
                {
                    string serverpath = HttpContext.Server.MapPath("~/img");
                    string filepath = serverpath + "/" + file.FileName;
                    if (file.FileName.Equals(""))
                    {
                        dichvu.Img = "";
                    }
                    else
                    {
                        file.SaveAs(filepath);
                        dichvu.Img = file.FileName;
                    }

                }
                dbcontext.dichvus.InsertOnSubmit(dichvu);
                dbcontext.SubmitChanges();
                return RedirectToAction("Index");
            }
            return View();
        }



        public ActionResult Delete(int id)
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
            dichvu dichvu = dbcontext.dichvus.FirstOrDefault(row => row.Id == id);
            return View(dichvu);
        }
        [HttpPost]
        public ActionResult Delete(int id, dichvu dv)
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
            dichvu dichvu = dbcontext.dichvus.FirstOrDefault(row => row.Id == id);
            if (dichvu != null)
            {
                dbcontext.dichvus.DeleteOnSubmit(dichvu);
                dbcontext.SubmitChanges();
            }
            return RedirectToAction("Index");
        }
        // Export Excel 
        public ActionResult ExportToExcel()
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();

            List<dichvu> all = dbcontext.dichvus.ToList();

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Students");
                var currentRow = 1;
                worksheet.Cell(currentRow, 1).Value = "Id";
                worksheet.Cell(currentRow, 2).Value = "Mã dịch vụ";
                worksheet.Cell(currentRow, 3).Value = "Tên dịch ụ";
                worksheet.Cell(currentRow, 4).Value = "Hình ảnh";
                worksheet.Cell(currentRow, 5).Value = "Mô tả";

                foreach (var nv in all)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = nv.Id;
                    worksheet.Cell(currentRow, 2).Value = nv.Madichvu;
                    worksheet.Cell(currentRow, 3).Value = nv.Tendichvu;
                    worksheet.Cell(currentRow, 4).Value = nv.Img;
                    worksheet.Cell(currentRow, 5).Value = nv.Mota;
                }
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();
                    return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Danhsachdichvu.xlsx");
                }
            }
        }
    }
}