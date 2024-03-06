using HouseCareWeb.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Razor.Generator;
using System.Web.WebPages;
using ClosedXML.Excel;
using System.IO;
using DocumentFormat.OpenXml.Office2010.Excel;

namespace HouseCareWeb.Controllers
{
    [Authorize(Users = "admin, nhanvien")]
    public class QuanLyNhanVienController : Controller
    {
        // GET: QuanLyNhanVien
        public ActionResult Index(String search="", String sortColumn = "Id", String iconClass = "fa-solid fa-sort-up",int page=1)
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
            
            // Find

            List<nhanvien> nhanvienList = dbcontext.nhanviens.Where(row => row.Hoten.Contains(search)).ToList();
            ViewBag.search = search;

            // Sort

            ViewBag.sortColumn = sortColumn;
            ViewBag.iconClass = iconClass;

            if (sortColumn == "Id")
            {
                if (iconClass == "fa-solid fa-sort-up")
                {
                    nhanvienList = nhanvienList.OrderBy(row => row.Id).ToList();
                }
                else
                {
                    nhanvienList = nhanvienList.OrderByDescending(row => row.Id).ToList();
                }
            }
            else if (sortColumn == "Manhanvien")
            {
                if (iconClass == "fa-solid fa-sort-up")
                {
                    nhanvienList = nhanvienList.OrderBy(row => row.Manhanvien).ToList();
                }
                else
                {
                    nhanvienList = nhanvienList.OrderByDescending(row => row.Manhanvien).ToList();
                }
            }
            else if (sortColumn == "Hoten")
            {
                if (iconClass == "fa-solid fa-sort-up")
                {
                    nhanvienList = nhanvienList.OrderBy(row => row.Hoten).ToList();
                }
                else
                {
                    nhanvienList = nhanvienList.OrderByDescending(row => row.Hoten).ToList();
                }
            }

            // Page

            int NoOfRecordPerPage = 5;
            int NoOfPages = Convert.ToInt32(Math.Ceiling(Convert.ToDouble(nhanvienList.Count)/Convert.ToDouble(NoOfRecordPerPage)));
            int NoOfRecordToSkip = (page - 1) * NoOfRecordPerPage;

            ViewBag.Page = page;
            ViewBag.NoOfPages = NoOfPages;

            nhanvienList = nhanvienList.Skip(NoOfRecordToSkip).Take(NoOfRecordPerPage).ToList();

            return View(nhanvienList);
        }
        public ActionResult Details(int id)
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
            nhanvien nhanvien = dbcontext.nhanviens.FirstOrDefault(row => row.Id == id);
            return View(nhanvien);
        }
        public ActionResult Edit(int id)
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
            nhanvien nhanvien = dbcontext.nhanviens.FirstOrDefault(row => row.Id == id);

            string preimg = nhanvien.Img;

            DateTime predate = nhanvien.Ngaysinh.Value;

            if (Request.Form.Count > 0)
            {
                String ten = Request.Form["Hoten"];
                String gioitinh = Request.Form["Gioitinh"];
                String diachi = Request.Form["Diachi"];
                String quequan = Request.Form["Quequan"];
                String chucvu = Request.Form["ChucVu"];
                int tuoi = int.Parse(Request.Form["Tuoi"]);
                try
                {
                    DateTime ngaysinh = DateTime.Parse(Request.Form["Ngaysinh"]);
                    nhanvien.Ngaysinh = ngaysinh;
                }
                catch
                {
                    nhanvien.Ngaysinh = predate;
                }
                
                String cccd = Request.Form["CCCD"];
                String email = Request.Form["Email"];
                String sdt = Request.Form["Sdt"];
                String mota = Request.Form["Mota"];
                String kinhnghiem = Request.Form["Kinhnghiem"];
                
                nhanvien.Hoten = ten;
                nhanvien.Gioitinh = gioitinh;
                nhanvien.Diachi = diachi;
                nhanvien.Quequan = quequan;
                nhanvien.Chucvu = chucvu;
                nhanvien.Tuoi = tuoi;
                nhanvien.CCCD = cccd;
                nhanvien.Email = email;
                nhanvien.Sdt = sdt;
                nhanvien.Mota = mota;
                nhanvien.Kinhnghiem = kinhnghiem;

                HttpPostedFileBase file = Request.Files["Img"];

                if (file != null)
                {
                    string serverpath = HttpContext.Server.MapPath("~/img");
                    string filepath = serverpath + "/" + file.FileName;
                    if (file.FileName.Equals(""))
                    {
                        nhanvien.Img = preimg;
                    }
                    else
                    {
                        file.SaveAs(filepath);
                        nhanvien.Img = file.FileName;
                    }
                }
                dbcontext.SubmitChanges();
                return RedirectToAction("Index");
            }
            return View(nhanvien);
        }

        public ActionResult Create()
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
            if (Request.Form.Count > 0)
            {
                String mnv = Request.Form["Manhanvien"];
                String ten = Request.Form["Hoten"];
                String gioitinh = Request.Form["Gioitinh"];
                String diachi = Request.Form["Diachi"];
                String quequan = Request.Form["Quequan"];
                String chucvu = Request.Form["ChucVu"];
                int tuoi = int.Parse(Request.Form["Tuoi"]);
                DateTime ngaysinh = DateTime.Parse(Request.Form["Ngaysinh"]);


                String cccd = Request.Form["CCCD"];
                String email = Request.Form["Email"];
                String sdt = Request.Form["Sdt"];
                String mota = Request.Form["Mota"];
                String kinhnghiem = Request.Form["Kinhnghiem"];

                nhanvien nhanvien = new nhanvien();

                nhanvien.Manhanvien = "NV"+mnv;
                nhanvien.Hoten = ten;
                nhanvien.Gioitinh = gioitinh;
                nhanvien.Ngaysinh = ngaysinh;
                nhanvien.Diachi = diachi;
                nhanvien.Quequan = quequan;
                nhanvien.Chucvu = chucvu;
                nhanvien.Tuoi = tuoi;
                nhanvien.CCCD = cccd;
                nhanvien.Email = email;
                nhanvien.Sdt = sdt;
                nhanvien.Mota = mota;
                nhanvien.Kinhnghiem = kinhnghiem;

                HttpPostedFileBase file = Request.Files["Img"];

                if (file != null)
                {
                    string serverpath = HttpContext.Server.MapPath("~/img");
                    string filepath = serverpath + "/" + file.FileName;
                    if (file.FileName.Equals(""))
                    {
                        nhanvien.Img = "";
                    }
                    else
                    {
                        file.SaveAs(filepath);
                        nhanvien.Img = file.FileName;
                    }
                }
                dbcontext.nhanviens.InsertOnSubmit(nhanvien);
                dbcontext.SubmitChanges();
                return RedirectToAction("Index");
            }
            return View();
        }
        public ActionResult Delete(int id)
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
            nhanvien nhanvien = dbcontext.nhanviens.FirstOrDefault(row => row.Id == id);
            return View(nhanvien);
        }
        [HttpPost]
        public ActionResult Delete(int id, nhanvien nv)
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
            nhanvien nhanvien = dbcontext.nhanviens.FirstOrDefault(row => row.Id == id);
            if (nhanvien != null)
            {
                dbcontext.nhanviens.DeleteOnSubmit(nhanvien);
                dbcontext.SubmitChanges();
                return RedirectToAction("Index");
            }
            return View();
        }

        // Export Excel 
        public ActionResult ExportToExcel()
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();

            List<nhanvien> all = dbcontext.nhanviens.ToList();

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Nhân viên");
                var currentRow = 1;
                worksheet.Cell(currentRow, 1).Value = "Id";
                worksheet.Cell(currentRow, 2).Value = "Mã nhân viên";
                worksheet.Cell(currentRow, 3).Value = "Họ tên";
                worksheet.Cell(currentRow, 4).Value = "Giới tính";
                worksheet.Cell(currentRow, 5).Value = "Địa chỉ";
                worksheet.Cell(currentRow, 6).Value = "Quê quán";
                worksheet.Cell(currentRow, 7).Value = "Chức vụ";
                worksheet.Cell(currentRow, 8).Value = "Hình ảnh";
                worksheet.Cell(currentRow, 9).Value = "Tuổi";
                worksheet.Cell(currentRow, 10).Value = "Ngày sinh";
                worksheet.Cell(currentRow, 11).Value = "CCCD/CMND";
                worksheet.Cell(currentRow, 12).Value = "Số điện thoại";
                worksheet.Cell(currentRow, 13).Value = "Email";
                worksheet.Cell(currentRow, 14).Value = "Mô tả";
                worksheet.Cell(currentRow, 15).Value = "Kinh nghiệm";

                foreach (var nv in all)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = nv.Id;
                    worksheet.Cell(currentRow, 2).Value = nv.Manhanvien;
                    worksheet.Cell(currentRow, 3).Value = nv.Hoten;
                    worksheet.Cell(currentRow, 4).Value = nv.Gioitinh;
                    worksheet.Cell(currentRow, 5).Value = nv.Diachi;
                    worksheet.Cell(currentRow, 6).Value = nv.Quequan;
                    worksheet.Cell(currentRow, 7).Value = nv.Chucvu;
                    worksheet.Cell(currentRow, 8).Value = nv.Img;
                    worksheet.Cell(currentRow, 9).Value = nv.Tuoi;
                    worksheet.Cell(currentRow, 10).Value = nv.Ngaysinh.Value.ToString("dd/MM/yyyy");
                    worksheet.Cell(currentRow, 11).Value = nv.CCCD;
                    worksheet.Cell(currentRow, 12).Value = nv.Sdt;
                    worksheet.Cell(currentRow, 13).Value = nv.Email;
                    worksheet.Cell(currentRow, 14).Value = nv.Mota;
                    worksheet.Cell(currentRow, 15).Value = nv.Kinhnghiem;


                }
                ////
                //var worksheet1 = workbook.Worksheets.Add("test");
                //var currentRow1 = 1;
                //worksheet1.Cell(currentRow1, 1).Value = "Id";
                //worksheet1.Cell(currentRow1, 2).Value = "Mã nhân viên";
                //worksheet1.Cell(currentRow1, 3).Value = "Họ tên";
                //worksheet1.Cell(currentRow1, 4).Value = "Giới tính";
                //worksheet1.Cell(currentRow1, 5).Value = "Địa chỉ";
                //worksheet1.Cell(currentRow1, 6).Value = "Quê quán";
                //worksheet1.Cell(currentRow1, 7).Value = "Chức vụ";
                //worksheet1.Cell(currentRow1, 8).Value = "Hình ảnh";
                //worksheet1.Cell(currentRow1, 9).Value = "Tuổi";
                //worksheet1.Cell(currentRow1, 10).Value = "Ngày sinh";
                //worksheet1.Cell(currentRow1, 11).Value = "CCCD/CMND";
                //worksheet1.Cell(currentRow1, 12).Value = "Số điện thoại";
                //worksheet1.Cell(currentRow1, 13).Value = "Email";
                //worksheet1.Cell(currentRow1, 14).Value = "Mô tả";
                //worksheet1.Cell(currentRow1, 15).Value = "Kinh nghiệm";

                //foreach (var nv in all)
                //{
                //    currentRow1++;
                //    worksheet1.Cell(currentRow1, 1).Value = nv.Id;
                //    worksheet1.Cell(currentRow1, 2).Value = nv.Manhanvien;
                //    worksheet1.Cell(currentRow1, 3).Value = nv.Hoten;
                //    worksheet1.Cell(currentRow1, 4).Value = nv.Gioitinh;
                //    worksheet1.Cell(currentRow1, 5).Value = nv.Diachi;
                //    worksheet1.Cell(currentRow1, 6).Value = nv.Quequan;
                //    worksheet1.Cell(currentRow1, 7).Value = nv.Chucvu;
                //    worksheet1.Cell(currentRow1, 8).Value = nv.Img;
                //    worksheet1.Cell(currentRow1, 9).Value = nv.Tuoi;
                //    worksheet1.Cell(currentRow1, 10).Value = nv.Ngaysinh.Value.ToString("dd/MM/yyyy");
                //    worksheet1.Cell(currentRow1, 11).Value = nv.CCCD;
                //    worksheet1.Cell(currentRow1, 12).Value = nv.Sdt;
                //    worksheet1.Cell(currentRow1, 13).Value = nv.Email;
                //    worksheet1.Cell(currentRow1, 14).Value = nv.Mota;
                //    worksheet1.Cell(currentRow1, 15).Value = nv.Kinhnghiem;

                //}
                ////
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();
                    return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Danhsachnhanvien.xlsx");
                }
            }
        }
    }
}