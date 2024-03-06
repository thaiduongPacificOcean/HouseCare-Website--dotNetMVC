using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using HouseCareWeb.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics.Eventing.Reader;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace HouseCareWeb.Controllers
{
    [Authorize(Users = "admin, nhanvien")]
    public class QuanLyDangKyGoiDichVuController : Controller
    {
        // GET: QuanLyDangKyGoiDichVu
        public ActionResult Index(String sortColumn = "Id", String iconClass = "fa-solid fa-sort-up", int page = 1)
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();

            List<dangkygoidichvu> listdk = dbcontext.dangkygoidichvus.ToList();

            // Lọc theo ngày 
            
            if (Request.Form.Count > 0)
            {
                try
                {
                    DateTime dateTo = DateTime.Parse(Request.Form["dateTo"]);
                    DateTime dateFrom = DateTime.Parse(Request.Form["dateFrom"]);
                    listdk = dbcontext.dangkygoidichvus.Where(x => x.Ngaydangky >= dateFrom && x.Ngaydangky <= dateTo).ToList();

                }
                catch
                {
                    listdk = dbcontext.dangkygoidichvus.ToList();
                }
            }

            // Sort

            ViewBag.sortColumn = sortColumn;
            ViewBag.iconClass = iconClass;

            if (sortColumn == "Id")
            {
                if (iconClass == "fa-solid fa-sort-up")
                {
                    listdk = listdk.OrderBy(row => row.Id).ToList();
                }
                else
                {
                    listdk = listdk.OrderByDescending(row => row.Id).ToList();
                }
            }
            else if (sortColumn == "Idtaikhoan")
            {
                if (iconClass == "fa-solid fa-sort-up")
                {
                    listdk = listdk.OrderBy(row => row.Idtaikhoan).ToList();
                }
                else
                {
                    listdk = listdk.OrderByDescending(row => row.Idtaikhoan).ToList();
                }
            }
            

            // Page

            int NoOfRecordPerPage = 5;
            int NoOfPages = Convert.ToInt32(Math.Ceiling(Convert.ToDouble(listdk.Count) / Convert.ToDouble(NoOfRecordPerPage)));
            int NoOfRecordToSkip = (page - 1) * NoOfRecordPerPage;

            ViewBag.Page = page;
            ViewBag.NoOfPages = NoOfPages;

            listdk = listdk.Skip(NoOfRecordToSkip).Take(NoOfRecordPerPage).ToList();

            return View(listdk);
        }

        // Danh sách đăng ký chưa duyệt
        public ActionResult Danhsachchuaduyet(String sortColumn = "Id", String iconClass = "fa-solid fa-sort-up", int page = 1)
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();

            List<dangkygoidichvu> listdk = dbcontext.dangkygoidichvus.Where(row => row.Trangthai == false).ToList();
            ViewBag.danhsachdangkygoidichvuchuaduyet = listdk;
            
            // Sort

            ViewBag.sortColumn = sortColumn;
            ViewBag.iconClass = iconClass;

            if (sortColumn == "Id")
            {
                if (iconClass == "fa-solid fa-sort-up")
                {
                    listdk = listdk.OrderBy(row => row.Id).ToList();
                }
                else
                {
                    listdk = listdk.OrderByDescending(row => row.Id).ToList();
                }
            }
            else if (sortColumn == "Idtaikhoan")
            {
                if (iconClass == "fa-solid fa-sort-up")
                {
                    listdk = listdk.OrderBy(row => row.Idtaikhoan).ToList();
                }
                else
                {
                    listdk = listdk.OrderByDescending(row => row.Idtaikhoan).ToList();
                }
            }


            // Page

            int NoOfRecordPerPage = 5;
            int NoOfPages = Convert.ToInt32(Math.Ceiling(Convert.ToDouble(listdk.Count) / Convert.ToDouble(NoOfRecordPerPage)));
            int NoOfRecordToSkip = (page - 1) * NoOfRecordPerPage;

            ViewBag.Page = page;
            ViewBag.NoOfPages = NoOfPages;

            listdk = listdk.Skip(NoOfRecordToSkip).Take(NoOfRecordPerPage).ToList();

            return View(listdk);
        }
        // Danhh sách đăng ký đã duyệt
        public ActionResult Danhsachdaduyet(String sortColumn = "Id", String iconClass = "fa-solid fa-sort-up", int page = 1)
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
            List<dangkygoidichvu> listdk = dbcontext.dangkygoidichvus.Where(row => row.Trangthai == true).ToList();
            ViewBag.danhsachdangkygoidichvudaduyet = listdk;

            // Sort

            ViewBag.sortColumn = sortColumn;
            ViewBag.iconClass = iconClass;

            if (sortColumn == "Id")
            {
                if (iconClass == "fa-solid fa-sort-up")
                {
                    listdk = listdk.OrderBy(row => row.Id).ToList();
                }
                else
                {
                    listdk = listdk.OrderByDescending(row => row.Id).ToList();
                }
            }
            else if (sortColumn == "Idtaikhoan")
            {
                if (iconClass == "fa-solid fa-sort-up")
                {
                    listdk = listdk.OrderBy(row => row.Idtaikhoan).ToList();
                }
                else
                {
                    listdk = listdk.OrderByDescending(row => row.Idtaikhoan).ToList();
                }
            }


            // Page

            int NoOfRecordPerPage = 5;
            int NoOfPages = Convert.ToInt32(Math.Ceiling(Convert.ToDouble(listdk.Count) / Convert.ToDouble(NoOfRecordPerPage)));
            int NoOfRecordToSkip = (page - 1) * NoOfRecordPerPage;

            ViewBag.Page = page;
            ViewBag.NoOfPages = NoOfPages;
            listdk = listdk.Skip(NoOfRecordToSkip).Take(NoOfRecordPerPage).ToList();

            return View(listdk);
        }
        // Danh sách đăng ký đã thực hiện
        public ActionResult Danhsachdathuchien(String sortColumn = "Id", String iconClass = "fa-solid fa-sort-up", int page = 1)
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();

            List<dangkygoidichvu> listdk = dbcontext.dangkygoidichvus.Where(row => row.Active == true).ToList();

            ViewBag.danhsachdangkydathuchien = listdk;

            // Sort

            ViewBag.sortColumn = sortColumn;
            ViewBag.iconClass = iconClass;

            if (sortColumn == "Id")
            {
                if (iconClass == "fa-solid fa-sort-up")
                {
                    listdk = listdk.OrderBy(row => row.Id).ToList();
                }
                else
                {
                    listdk = listdk.OrderByDescending(row => row.Id).ToList();
                }
            }
            else if (sortColumn == "Idtaikhoan")
            {
                if (iconClass == "fa-solid fa-sort-up")
                {
                    listdk = listdk.OrderBy(row => row.Idtaikhoan).ToList();
                }
                else
                {
                    listdk = listdk.OrderByDescending(row => row.Idtaikhoan).ToList();
                }
            }


            // Page

            int NoOfRecordPerPage = 5;
            int NoOfPages = Convert.ToInt32(Math.Ceiling(Convert.ToDouble(listdk.Count) / Convert.ToDouble(NoOfRecordPerPage)));
            int NoOfRecordToSkip = (page - 1) * NoOfRecordPerPage;

            ViewBag.Page = page;
            ViewBag.NoOfPages = NoOfPages;
            listdk = listdk.Skip(NoOfRecordToSkip).Take(NoOfRecordPerPage).ToList();

            return View(listdk);

        }
        // Danh sách đăng ký chưa thực hiện
        public ActionResult Danhsachchuathuchien(String sortColumn = "Id", String iconClass = "fa-solid fa-sort-up", int page = 1)
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();

            List<dangkygoidichvu> listdk = dbcontext.dangkygoidichvus.Where(row => row.Active == false).ToList();

            ViewBag.danhsachdangkychuathuchien = listdk;

            // Sort

            ViewBag.sortColumn = sortColumn;
            ViewBag.iconClass = iconClass;

            if (sortColumn == "Id")
            {
                if (iconClass == "fa-solid fa-sort-up")
                {
                    listdk = listdk.OrderBy(row => row.Id).ToList();
                }
                else
                {
                    listdk = listdk.OrderByDescending(row => row.Id).ToList();
                }
            }
            else if (sortColumn == "Idtaikhoan")
            {
                if (iconClass == "fa-solid fa-sort-up")
                {
                    listdk = listdk.OrderBy(row => row.Idtaikhoan).ToList();
                }
                else
                {
                    listdk = listdk.OrderByDescending(row => row.Idtaikhoan).ToList();
                }
            }


            // Page

            int NoOfRecordPerPage = 5;
            int NoOfPages = Convert.ToInt32(Math.Ceiling(Convert.ToDouble(listdk.Count) / Convert.ToDouble(NoOfRecordPerPage)));
            int NoOfRecordToSkip = (page - 1) * NoOfRecordPerPage;

            ViewBag.Page = page;
            ViewBag.NoOfPages = NoOfPages;
            listdk = listdk.Skip(NoOfRecordToSkip).Take(NoOfRecordPerPage).ToList();

            return View(listdk);

        }
        public ActionResult Details(int id)
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
            dangkygoidichvu dangkygoidichvu = dbcontext.dangkygoidichvus.FirstOrDefault(row => row.Id == id);
            return View(dangkygoidichvu);
        }
        
        public ActionResult Edit(int id)
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
            dangkygoidichvu dangkygoidichvu = dbcontext.dangkygoidichvus.FirstOrDefault(row => row.Id == id);
            List<dangkygoidichvu> dkall = dbcontext.dangkygoidichvus.ToList();
            ViewBag.listall = dkall;
            List<account> accounts = dbcontext.accounts.ToList();
            List<goidichvu> goidv = dbcontext.goidichvus.ToList();
            ViewBag.goidv = goidv;
            ViewBag.acc = accounts;

            int premagoi = dangkygoidichvu.Magoi.Value;
            DateTime predate = dangkygoidichvu.Ngaydangky.Value;
            bool prestate = dangkygoidichvu.Trangthai.Value;
            bool prestaeA = dangkygoidichvu.Active.Value;
            if (Request.Form.Count > 0)
            {
                try
                {
                    int magoi = int.Parse(Request.Form["Magoi"]);
                    dangkygoidichvu.Magoi = magoi;

                }
                catch
                {
                    dangkygoidichvu.Magoi = premagoi;
                }

                try
                {
                    DateTime date = DateTime.Parse(Request.Form["Ngaydangky"]);
                    dangkygoidichvu.Ngaydangky = date;
                }
                catch
                {
                    dangkygoidichvu.Ngaydangky = predate;
                }
                String diachi = Request.Form["Diachi"];
                String yeucau = Request.Form["Yeucau"];
                try
                {
                    bool trangthai = bool.Parse(Request.Form["Trangthai"]);
                    dangkygoidichvu.Trangthai = trangthai;

                }
                catch
                {
                    dangkygoidichvu.Trangthai = prestate;
                }
                try
                {
                    bool active = bool.Parse(Request.Form["Active"]);
                    dangkygoidichvu.Active = active;

                }
                catch
                {
                    dangkygoidichvu.Active = prestaeA;
                }
                dangkygoidichvu.Diachi = diachi;
                dangkygoidichvu.Yeucau= yeucau;

                dbcontext.SubmitChanges();
                return RedirectToAction("Index");
            }
            return View(dangkygoidichvu);
        }
        public ActionResult Delete(int id)
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
            dangkygoidichvu dangkygoidichvu = dbcontext.dangkygoidichvus.FirstOrDefault(row => row.Id == id);
            
            return View(dangkygoidichvu);
        }
        [HttpPost]
        public ActionResult Delete(int id, dangkygoidichvu dk)
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
            dangkygoidichvu dangkygoidichvu = dbcontext.dangkygoidichvus.FirstOrDefault(row => row.Id == id);
            if (dangkygoidichvu != null)
            {
                dbcontext.dangkygoidichvus.DeleteOnSubmit(dangkygoidichvu);
                dbcontext.SubmitChanges();
                return RedirectToAction("Index");
            }
            return View();
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
            return RedirectToAction("Index");
        }
        // Hoan thanh 
        public ActionResult Active(int id)
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();

            dangkygoidichvu dangkygoidichvu = dbcontext.dangkygoidichvus.FirstOrDefault(row => row.Id == id);

            if (dangkygoidichvu != null)
            {
                dangkygoidichvu.Active = true;

                dbcontext.SubmitChanges();
            }
            return RedirectToAction("Index");


        }
        // Export Excel Danh sách all
        public ActionResult ExportToExcel()
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();

            List<dangkygoidichvu> all = dbcontext.dangkygoidichvus.ToList();
            List<dangkygoidichvu> allchuaduyet = dbcontext.dangkygoidichvus.Where(row=>row.Trangthai==false).ToList();
            List<dangkygoidichvu> alldaduyet = dbcontext.dangkygoidichvus.Where(row => row.Trangthai == true).ToList();
            List<dangkygoidichvu> allchuathuchien = dbcontext.dangkygoidichvus.Where(row => row.Active == false).ToList();
            List<dangkygoidichvu> alldathuchien = dbcontext.dangkygoidichvus.Where(row => row.Active == true).ToList();

            List<account> acc = dbcontext.accounts.ToList();
            List<goidichvu> goidichvu = dbcontext.goidichvus.ToList();


            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Danh sách đăng ký gói dịch vụ");
                var worksheet2 = workbook.Worksheets.Add("DSDK chưa duyệt");
                var worksheet3 = workbook.Worksheets.Add("DSDK đã duyệt");
                var worksheet4 = workbook.Worksheets.Add("DSDK chưa thực hiện");
                var worksheet5 = workbook.Worksheets.Add("DSDK đã thực hiện");

                var currentRow = 1;
                worksheet.Cell(currentRow, 1).Value = "Id";
                worksheet.Cell(currentRow, 2).Value = "Tên tài khoản đăng ký";
                worksheet.Cell(currentRow, 3).Value = "Mã tài khoản";
                worksheet.Cell(currentRow, 4).Value = "Mã gói";
                worksheet.Cell(currentRow, 5).Value = "Tên gói";
                worksheet.Cell(currentRow, 6).Value = "Ngày đăng ký";
                worksheet.Cell(currentRow, 7).Value = "Giờ";
                worksheet.Cell(currentRow, 8).Value = "Địa chỉ";
                worksheet.Cell(currentRow, 9).Value = "Yêu cầu";
                worksheet.Cell(currentRow, 10).Value = "Trạng thái xác nhận";
                worksheet.Cell(currentRow, 11).Value = "Thời lượng";
                worksheet.Cell(currentRow, 12).Value = "Trạng thái thực hiện";
                foreach (var nv in all)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = nv.Id;
                    foreach (var accdk in acc)
                    {
                        if (accdk.Id == nv.Idtaikhoan)
                        {
                            worksheet.Cell(currentRow,2).Value = accdk.Ten;
                            worksheet.Cell(currentRow, 3).Value = accdk.Mataikhoan;

                        }
                    }
                    //worksheet.Cell(currentRow, 4).Value = nv.Idtaikhoan;
                    worksheet.Cell(currentRow, 4).Value = nv.Magoi;
                    foreach (var goi in goidichvu)
                    {
                        if (goi.Id == nv.Magoi)
                        {
                            worksheet.Cell(currentRow, 5).Value = goi.Tengoi;

                        }
                    }
                    worksheet.Cell(currentRow, 6).Value = nv.Ngaydangky.Value.ToString("dd/MM/yyyy");
                    worksheet.Cell(currentRow, 7).Value = nv.Gio;
                    worksheet.Cell(currentRow, 8).Value = nv.Diachi;
                    worksheet.Cell(currentRow, 9).Value = nv.Yeucau;
                    if (nv.Trangthai == true)
                    {
                        worksheet.Cell(currentRow, 10).Value = "Đã duyệt";

                    }
                    else
                    {
                        worksheet.Cell(currentRow, 10).Value = "Chưa duyệt";

                    }
                    worksheet.Cell(currentRow, 11).Value = "1 giờ";

                    if (nv.Active == true)
                    {
                        worksheet.Cell(currentRow, 12).Value = "Đã hoàn thành";

                    }
                    else
                    {
                        worksheet.Cell(currentRow, 12).Value = "Chưa hoàn thành";

                    }
                    
                }

                // danh sách chưa duyệt

                var currentRowws2 = 1;
                worksheet2.Cell(currentRowws2, 1).Value = "Id";
                worksheet2.Cell(currentRowws2, 2).Value = "Tên tài khoản đăng ký";
                worksheet2.Cell(currentRowws2, 3).Value = "Mã tài khoản";
                worksheet2.Cell(currentRowws2, 4).Value = "Mã gói";
                worksheet2.Cell(currentRowws2, 5).Value = "Ngày đăng ký";
                worksheet2.Cell(currentRowws2, 6).Value = "Giờ";
                worksheet2.Cell(currentRowws2, 7).Value = "Địa chỉ";
                worksheet2.Cell(currentRowws2, 8).Value = "Yêu cầu";
                worksheet2.Cell(currentRowws2, 9).Value = "Trạng thái xác nhận";
                worksheet2.Cell(currentRowws2, 10).Value = "Thời lượng";
                foreach (var nv in allchuaduyet)
                {
                    currentRowws2++;
                    worksheet2.Cell(currentRowws2, 1).Value = nv.Id;
                    foreach (var accdk in acc)
                    {
                        if (accdk.Id == nv.Idtaikhoan)
                        {
                            worksheet2.Cell(currentRowws2, 2).Value = accdk.Ten;

                        }
                    }
                    worksheet2.Cell(currentRowws2, 3).Value = nv.Idtaikhoan;
                    worksheet2.Cell(currentRowws2, 4).Value = nv.Magoi;
                    worksheet2.Cell(currentRowws2, 5).Value = nv.Ngaydangky.Value.ToString("dd/MM/yyyy");
                    worksheet2.Cell(currentRowws2, 6).Value = nv.Gio;
                    worksheet2.Cell(currentRowws2, 7).Value = nv.Diachi;
                    worksheet2.Cell(currentRowws2, 8).Value = nv.Yeucau;
                    if (nv.Trangthai == true)
                    {
                        worksheet2.Cell(currentRowws2, 9).Value = "Đã duyệt";

                    }
                    else
                    {
                        worksheet2.Cell(currentRowws2, 9).Value = "Chưa duyệt";

                    }
                    worksheet2.Cell(currentRowws2, 10).Value = "1 giờ";

                    

                }
                // danh sách đã duyệt
                var currentRow3 = 1;
                worksheet.Cell(currentRow3, 1).Value = "Id";
                worksheet.Cell(currentRow3, 2).Value = "Tên tài khoản đăng ký";
                worksheet.Cell(currentRow3, 3).Value = "Mã tài khoản";
                worksheet.Cell(currentRow3, 4).Value = "Mã gói";
                worksheet.Cell(currentRow3, 5).Value = "Tên gói";
                worksheet.Cell(currentRow3, 6).Value = "Ngày đăng ký";
                worksheet.Cell(currentRow3, 7).Value = "Giờ";
                worksheet.Cell(currentRow3, 8).Value = "Địa chỉ";
                worksheet.Cell(currentRow3, 9).Value = "Yêu cầu";
                worksheet.Cell(currentRow3, 10).Value = "Trạng thái xác nhận";
                worksheet.Cell(currentRow3, 11).Value = "Thời lượng";
                worksheet.Cell(currentRow3, 12).Value = "Trạng thái thực hiện";
                foreach (var nv in alldaduyet)
                {
                    currentRow3++;
                    worksheet3.Cell(currentRow3, 1).Value = nv.Id;
                    foreach (var accdk in acc)
                    {
                        if (accdk.Id == nv.Idtaikhoan)
                        {
                            worksheet3.Cell(currentRow3, 2).Value = accdk.Ten;
                            worksheet3.Cell(currentRow3, 3).Value = accdk.Mataikhoan;

                        }
                    }
                    //worksheet.Cell(currentRow, 4).Value = nv.Idtaikhoan;
                    worksheet3.Cell(currentRow3, 4).Value = nv.Magoi;
                    foreach (var goi in goidichvu)
                    {
                        if (goi.Id == nv.Magoi)
                        {
                            worksheet3.Cell(currentRow3, 5).Value = goi.Tengoi;

                        }
                    }
                    worksheet3.Cell(currentRow3, 6).Value = nv.Ngaydangky.Value.ToString("dd/MM/yyyy");
                    worksheet3.Cell(currentRow3, 7).Value = nv.Gio;
                    worksheet3.Cell(currentRow3, 8).Value = nv.Diachi;
                    worksheet3.Cell(currentRow3, 9).Value = nv.Yeucau;
                    if (nv.Trangthai == true)
                    {
                        worksheet3.Cell(currentRow3, 10).Value = "Đã duyệt";
                    }
                    else
                    {
                        worksheet3.Cell(currentRow3, 10).Value = "Chưa duyệt";

                    }
                    worksheet3.Cell(currentRow3, 11).Value = "1 giờ";

                    if (nv.Active == true)
                    {
                        worksheet3.Cell(currentRow3, 12).Value = "Đã hoàn thành";

                    }
                    else
                    {
                        worksheet3.Cell(currentRow3, 12).Value = "Chưa hoàn thành";

                    }

                }

                // danh sách chưa thực hiện
                var currentRow4 = 1;
                worksheet4.Cell(currentRow4, 1).Value = "Id";
                worksheet4.Cell(currentRow4, 2).Value = "Tên tài khoản đăng ký";
                worksheet4.Cell(currentRow4, 3).Value = "Mã tài khoản";
                worksheet4.Cell(currentRow4, 4).Value = "Mã gói";
                worksheet4.Cell(currentRow4, 5).Value = "Tên gói";
                worksheet4.Cell(currentRow4, 6).Value = "Ngày đăng ký";
                worksheet4.Cell(currentRow4, 7).Value = "Giờ";
                worksheet4.Cell(currentRow4, 8).Value = "Địa chỉ";
                worksheet4.Cell(currentRow4, 9).Value = "Yêu cầu";
                worksheet4.Cell(currentRow4, 10).Value = "Trạng thái xác nhận";
                worksheet4.Cell(currentRow4, 11).Value = "Thời lượng";
                worksheet4.Cell(currentRow4, 12).Value = "Trạng thái thực hiện";
                foreach (var nv in allchuathuchien)
                {
                    currentRow4++;
                    worksheet4.Cell(currentRow4, 1).Value = nv.Id;
                    foreach (var accdk in acc)
                    {
                        if (accdk.Id == nv.Idtaikhoan)
                        {
                            worksheet4.Cell(currentRow4, 2).Value = accdk.Ten;
                            worksheet4.Cell(currentRow4, 3).Value = accdk.Mataikhoan;

                        }
                    }
                    //worksheet.Cell(currentRow, 4).Value = nv.Idtaikhoan;
                    worksheet.Cell(currentRow4, 4).Value = nv.Magoi;
                    foreach (var goi in goidichvu)
                    {
                        if (goi.Id == nv.Magoi)
                        {
                            worksheet4.Cell(currentRow4, 5).Value = goi.Tengoi;

                        }
                    }
                    worksheet4.Cell(currentRow4, 6).Value = nv.Ngaydangky.Value.ToString("dd/MM/yyyy");
                    worksheet4.Cell(currentRow4, 7).Value = nv.Gio;
                    worksheet4.Cell(currentRow4, 8).Value = nv.Diachi;
                    worksheet4.Cell(currentRow4, 9).Value = nv.Yeucau;
                    if (nv.Trangthai == true)
                    {
                        worksheet4.Cell(currentRow4, 10).Value = "Đã duyệt";

                    }
                    else
                    {
                        worksheet4.Cell(currentRow4, 10).Value = "Chưa duyệt";

                    }
                    worksheet4.Cell(currentRow4, 11).Value = "1 giờ";

                    if (nv.Active == true)
                    {
                        worksheet4.Cell(currentRow4, 12).Value = "Đã hoàn thành";

                    }
                    else
                    {
                        worksheet4.Cell(currentRow4, 12).Value = "Chưa hoàn thành";

                    }
                    
                    }
                // da thuc hien
                var currentRow5 = 1;
                worksheet5.Cell(currentRow5, 1).Value = "Id";
                worksheet5.Cell(currentRow5, 2).Value = "Tên tài khoản đăng ký";
                worksheet5.Cell(currentRow5, 3).Value = "Mã tài khoản";
                worksheet5.Cell(currentRow5, 4).Value = "Mã gói";
                worksheet5.Cell(currentRow5, 5).Value = "Tên gói";
                worksheet5.Cell(currentRow5, 6).Value = "Ngày đăng ký";
                worksheet5.Cell(currentRow5, 7).Value = "Giờ";
                worksheet5.Cell(currentRow5, 8).Value = "Địa chỉ";
                worksheet5.Cell(currentRow5, 9).Value = "Yêu cầu";
                worksheet5.Cell(currentRow5, 10).Value = "Trạng thái xác nhận";
                worksheet5.Cell(currentRow5, 11).Value = "Thời lượng";
                worksheet5.Cell(currentRow5, 12).Value = "Trạng thái thực hiện";
                foreach (var nv in alldathuchien)
                {
                    currentRow5++;
                    worksheet5.Cell(currentRow5, 1).Value = nv.Id;
                    foreach (var accdk in acc)
                    {
                        if (accdk.Id == nv.Idtaikhoan)
                        {
                            worksheet5.Cell(currentRow5, 2).Value = accdk.Ten;
                            worksheet5.Cell(currentRow5, 3).Value = accdk.Mataikhoan;

                        }
                    }
                    //worksheet.Cell(currentRow, 4).Value = nv.Idtaikhoan;
                    worksheet5.Cell(currentRow5, 4).Value = nv.Magoi;
                    foreach (var goi in goidichvu)
                    {
                        if (goi.Id == nv.Magoi)
                        {
                            worksheet5.Cell(currentRow5, 5).Value = goi.Tengoi;

                        }
                    }
                    worksheet5.Cell(currentRow5, 6).Value = nv.Ngaydangky.Value.ToString("dd/MM/yyyy");
                    worksheet5.Cell(currentRow5, 7).Value = nv.Gio;
                    worksheet5.Cell(currentRow5, 8).Value = nv.Diachi;
                    worksheet5.Cell(currentRow5, 9).Value = nv.Yeucau;
                    if (nv.Trangthai == true)
                    {
                        worksheet5.Cell(currentRow5, 10).Value = "Đã duyệt";

                    }
                    else
                    {
                        worksheet5.Cell(currentRow5, 10).Value = "Chưa duyệt";

                    }
                    worksheet5.Cell(currentRow5, 11).Value = "1 giờ";

                    if (nv.Active == true)
                    {
                        worksheet5.Cell(currentRow5, 12).Value = "Đã hoàn thành";

                    }
                    else
                    {
                        worksheet5.Cell(currentRow5, 12).Value = "Chưa hoàn thành";

                    }

                }
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();
                    return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Danhsachdangky.xlsx");
                }
            }
        }
        // Xuất danh sách chưa duyệt
        public ActionResult ExportToExcel2()
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();

            List<dangkygoidichvu> all = dbcontext.dangkygoidichvus.Where(row => row.Trangthai == false).ToList();

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Danh sách đăng ký gói dịch vụ");
                var currentRow = 1;
                worksheet.Cell(currentRow, 1).Value = "Id";
                worksheet.Cell(currentRow, 2).Value = "Mã đăng ký";
                worksheet.Cell(currentRow, 3).Value = "Mã tài khoản";
                worksheet.Cell(currentRow, 4).Value = "Mã gói";
                worksheet.Cell(currentRow, 5).Value = "Ngày đăng ký";
                worksheet.Cell(currentRow, 6).Value = "Giờ";
                worksheet.Cell(currentRow, 7).Value = "Địa chỉ";
                worksheet.Cell(currentRow, 8).Value = "Yêu cầu";
                worksheet.Cell(currentRow, 9).Value = "Trạng thái xác nhận";
                worksheet.Cell(currentRow, 10).Value = "Thời lượng";
                worksheet.Cell(currentRow, 11).Value = "Trạng thái thực hiện";

                foreach (var nv in all)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = nv.Id;
                    worksheet.Cell(currentRow, 2).Value = nv.Madangky;
                    worksheet.Cell(currentRow, 3).Value = nv.Idtaikhoan;
                    worksheet.Cell(currentRow, 4).Value = nv.Magoi;
                    worksheet.Cell(currentRow, 5).Value = nv.Ngaydangky.Value.ToString("dd/MM/yyyy");
                    worksheet.Cell(currentRow, 6).Value = nv.Gio;
                    worksheet.Cell(currentRow, 7).Value = nv.Diachi;
                    worksheet.Cell(currentRow, 8).Value = nv.Yeucau;
                    worksheet.Cell(currentRow, 9).Value = nv.Trangthai;
                    worksheet.Cell(currentRow, 10).Value = nv.Thoiluong;
                    worksheet.Cell(currentRow, 11).Value = nv.Active;


                }
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();
                    return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Danhsachdangkychuaduyet.xlsx");
                }
            }
        }
        // danh sách đăng ký đã duyệt
        public ActionResult ExportToExcel3()
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();

            List<dangkygoidichvu> all = dbcontext.dangkygoidichvus.Where(row => row.Trangthai == true).ToList();

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Danh sách đăng ký gói dịch vụ");
                var currentRow = 1;
                worksheet.Cell(currentRow, 1).Value = "Id";
                worksheet.Cell(currentRow, 2).Value = "Mã đăng ký";
                worksheet.Cell(currentRow, 3).Value = "Mã tài khoản";
                worksheet.Cell(currentRow, 4).Value = "Mã gói";
                worksheet.Cell(currentRow, 5).Value = "Ngày đăng ký";
                worksheet.Cell(currentRow, 6).Value = "Giờ";
                worksheet.Cell(currentRow, 7).Value = "Địa chỉ";
                worksheet.Cell(currentRow, 8).Value = "Yêu cầu";
                worksheet.Cell(currentRow, 9).Value = "Trạng thái xác nhận";
                worksheet.Cell(currentRow, 10).Value = "Thời lượng";
                worksheet.Cell(currentRow, 11).Value = "Trạng thái thực hiện";

                foreach (var nv in all)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = nv.Id;
                    worksheet.Cell(currentRow, 2).Value = nv.Madangky;
                    worksheet.Cell(currentRow, 3).Value = nv.Idtaikhoan;
                    worksheet.Cell(currentRow, 4).Value = nv.Magoi;
                    worksheet.Cell(currentRow, 5).Value = nv.Ngaydangky.Value.ToString("dd/MM/yyyy");
                    worksheet.Cell(currentRow, 6).Value = nv.Gio;
                    worksheet.Cell(currentRow, 7).Value = nv.Diachi;
                    worksheet.Cell(currentRow, 8).Value = nv.Yeucau;
                    worksheet.Cell(currentRow, 9).Value = nv.Trangthai;
                    worksheet.Cell(currentRow, 10).Value = nv.Thoiluong;
                    worksheet.Cell(currentRow, 11).Value = nv.Active;


                }
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();
                    return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Danhsachdangkydaduyet.xlsx");
                }
            }
        }
        // danh sach đăng ký đã thực hiện
        public ActionResult ExportToExcel4()
        {
            housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();

            List<dangkygoidichvu> all = dbcontext.dangkygoidichvus.Where(row => row.Active == true).ToList();

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Danh sách đăng ký gói dịch vụ");
                var currentRow = 1;
                worksheet.Cell(currentRow, 1).Value = "Id";
                worksheet.Cell(currentRow, 2).Value = "Mã đăng ký";
                worksheet.Cell(currentRow, 3).Value = "Mã tài khoản";
                worksheet.Cell(currentRow, 4).Value = "Mã gói";
                worksheet.Cell(currentRow, 5).Value = "Ngày đăng ký";
                worksheet.Cell(currentRow, 6).Value = "Giờ";
                worksheet.Cell(currentRow, 7).Value = "Địa chỉ";
                worksheet.Cell(currentRow, 8).Value = "Yêu cầu";
                worksheet.Cell(currentRow, 9).Value = "Trạng thái xác nhận";
                worksheet.Cell(currentRow, 10).Value = "Thời lượng";
                worksheet.Cell(currentRow, 11).Value = "Trạng thái thực hiện";

                foreach (var nv in all)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = nv.Id;
                    worksheet.Cell(currentRow, 2).Value = nv.Madangky;
                    worksheet.Cell(currentRow, 3).Value = nv.Idtaikhoan;
                    worksheet.Cell(currentRow, 4).Value = nv.Magoi;
                    worksheet.Cell(currentRow, 5).Value = nv.Ngaydangky.Value.ToString("dd/MM/yyyy");
                    worksheet.Cell(currentRow, 6).Value = nv.Gio;
                    worksheet.Cell(currentRow, 7).Value = nv.Diachi;
                    worksheet.Cell(currentRow, 8).Value = nv.Yeucau;
                    worksheet.Cell(currentRow, 9).Value = nv.Trangthai;
                    worksheet.Cell(currentRow, 10).Value = nv.Thoiluong;
                    worksheet.Cell(currentRow, 11).Value = nv.Active;


                }
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();
                    return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Danhsachdangkydathuchien.xlsx");
                }
            }
        }
    }
}