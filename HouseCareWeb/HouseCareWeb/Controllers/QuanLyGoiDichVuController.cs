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
    public class QuanLyGoiDichVuController : Controller
    {
        // GET: QuanLyGoiDichVu
        public ActionResult Index(String search="",String sortColumn = "Id", String iconClass = "fa-solid fa-sort-up", int page = 1)
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
            
            List<goidichvu> goidichvuList = dbcontext.goidichvus.Where(row => row.Tengoi.Contains(search)).ToList();
            ViewBag.search = search;

            // sort
            ViewBag.sortColumn = sortColumn;
            ViewBag.iconClass = iconClass;

            if (sortColumn == "Magoi")
            {
                if (iconClass == "fa-solid fa-sort-up")
                {
                    goidichvuList = goidichvuList.OrderBy(row => row.Magoi).ToList();
                }
                else
                {
                    goidichvuList = goidichvuList.OrderByDescending(row => row.Magoi).ToList();
                }
            }
            else if (sortColumn == "Tengoi")
            {
                if (iconClass == "fa-solid fa-sort-up")
                {
                    goidichvuList = goidichvuList.OrderBy(row => row.Tengoi).ToList();
                }
                else
                {
                    goidichvuList = goidichvuList.OrderByDescending(row => row.Tengoi).ToList();
                }
            }
            else if (sortColumn == "Giatien")
            {
                if (iconClass == "fa-solid fa-sort-up")
                {
                    goidichvuList = goidichvuList.OrderBy(row => row.Giatien).ToList();
                }
                else
                {
                    goidichvuList = goidichvuList.OrderByDescending(row => row.Giatien).ToList();
                }
            }
            else if (sortColumn == "Madichvu")
            {
                if (iconClass == "fa-solid fa-sort-up")
                {
                    goidichvuList = goidichvuList.OrderBy(row => row.Madichvu).ToList();
                }
                else
                {
                    goidichvuList = goidichvuList.OrderByDescending(row => row.Madichvu).ToList();
                }
            }
            // Page

            int NoOfRecordPerPage = 5;
            int NoOfPages = Convert.ToInt32(Math.Ceiling(Convert.ToDouble(goidichvuList.Count) / Convert.ToDouble(NoOfRecordPerPage)));
            int NoOfRecordToSkip = (page - 1) * NoOfRecordPerPage;

            ViewBag.Page = page;
            ViewBag.NoOfPages = NoOfPages;

            goidichvuList = goidichvuList.Skip(NoOfRecordToSkip).Take(NoOfRecordPerPage).ToList();


            return View(goidichvuList);
        }
        public ActionResult Details(int id)
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
            goidichvu goidichvu = dbcontext.goidichvus.FirstOrDefault(row => row.Id == id);
            return View(goidichvu);
        }
        public ActionResult Create()
        {
            housecareDBContextDataContext context = new housecareDBContextDataContext();
            List<dichvu> dichvu = context.dichvus.ToList();
            ViewBag.goidichvuList = dichvu;
            int premadichvu = 1;
            if (Request.Form.Count > 0)
            {
                goidichvu goi = new goidichvu();
                String magoi = Request.Form["Magoi"];

                String tengoi = Request.Form["Tengoi"];
                float giatien = float.Parse(Request.Form["Giatien"]);
                String detail = Request.Form["Detail"];
                String loai = Request.Form["Loai"];
                try
                {
                    int madichvu = int.Parse(Request.Form["Madichvu"]);
                    goi.Madichvu = madichvu;

                }
                catch
                {
                    goi.Madichvu = premadichvu;
                }
                goi.Tengoi = tengoi;
                goi.Giatien = giatien;
                goi.Detail = detail;
                goi.Loai = loai;
                goi.Magoi = "GOIDV"+ magoi;
                context.goidichvus.InsertOnSubmit(goi);
                context.SubmitChanges();
                return RedirectToAction("Index");
            }
            return View();
        }
        public ActionResult Edit(int id)
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
            goidichvu goi = dbcontext.goidichvus.FirstOrDefault(row => row.Id == id);

            List<dichvu> dichvu = dbcontext.dichvus.ToList();
            ViewBag.goidichvuList = dichvu;

            int premadichvu = goi.Madichvu.Value;

            if (Request.Form.Count > 0)
            {

                String tengoi = Request.Form["Tengoi"];
                float giatien = float.Parse(Request.Form["Giatien"]);
                String detail = Request.Form["Detail"];
                String loai = Request.Form["Loai"];
                try
                {
                    int madichvu = int.Parse(Request.Form["Madichvu"]);
                    goi.Madichvu = madichvu;

                }
                catch
                {
                    goi.Madichvu = premadichvu;
                }

                goi.Tengoi = tengoi;
                goi.Giatien=giatien;
                goi.Detail=detail;
                goi.Loai=loai;

                dbcontext.SubmitChanges();
                return RedirectToAction("Index");
            }
            return View(goi);
        }
        public ActionResult Delete(int id)
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
            goidichvu goi = dbcontext.goidichvus.FirstOrDefault(row => row.Id == id);
            return View(goi);
        }
        [HttpPost]
        public ActionResult Delete(int id , goidichvu g)
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
            goidichvu goi = dbcontext.goidichvus.FirstOrDefault(row => row.Id == id);
            if (goi != null)
            {
                dbcontext.goidichvus.DeleteOnSubmit(goi);
                dbcontext.SubmitChanges();
            }
            return RedirectToAction("Index");
        }
        // Export Excel 
        public ActionResult ExportToExcel()
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();

            List<goidichvu> all = dbcontext.goidichvus.ToList();

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Students");
                var currentRow = 1;
                worksheet.Cell(currentRow, 1).Value = "Id";
                worksheet.Cell(currentRow, 2).Value = "Mã gói";
                worksheet.Cell(currentRow, 3).Value = "Tên gói";
                worksheet.Cell(currentRow, 4).Value = "Giá tiền";
                worksheet.Cell(currentRow, 5).Value = "Mã dịch vụ";
                worksheet.Cell(currentRow, 6).Value = "Mô tả";
                worksheet.Cell(currentRow, 7).Value = "Loại";

                foreach (var nv in all)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = nv.Id;
                    worksheet.Cell(currentRow, 2).Value = nv.Magoi;
                    worksheet.Cell(currentRow, 3).Value = nv.Tengoi;
                    worksheet.Cell(currentRow, 4).Value = nv.Giatien;
                    worksheet.Cell(currentRow, 5).Value = nv.Madichvu;
                    worksheet.Cell(currentRow, 6).Value = nv.Detail;
                    worksheet.Cell(currentRow, 7).Value = nv.Loai;


                }
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();
                    return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Danhsachgoidichvu.xlsx");
                }
            }
        }
    }
}