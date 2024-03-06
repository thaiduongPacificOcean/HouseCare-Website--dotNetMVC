using HouseCareWeb.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Security;

namespace HouseCareWeb.Controllers
{
    public class LoginController : Controller
    {
        private housecareDBContextDataContext dbcontext = new housecareDBContextDataContext();
        // GET: Login
        public ActionResult Index()
        {
            return View();
        }
        [HttpPost]
        public ActionResult Login(account acc , String returnUrl)
        {
            if (ModelState.IsValid)
            {
                var us = dbcontext.accounts.Where(s => s.Taikhoan.Equals(acc.Taikhoan) && s.Matkhau.Equals(acc.Matkhau)).ToList();

                if (us.Count > 0)
                {
                    if (acc.Taikhoan == "admin" && acc.Matkhau == "admin")
                    {
                        FormsAuthentication.SetAuthCookie(acc.Taikhoan, false);
                        if (Url.IsLocalUrl(returnUrl) && returnUrl.Length > 1 && returnUrl.StartsWith("/")
                            && !returnUrl.StartsWith("//") && !returnUrl.StartsWith("/\\"))
                        {
                            return Redirect(returnUrl);
                        }
                        else
                        {
                            Session["Taikhoan"] = us.FirstOrDefault(x => x.Id == us.First().Id);
                            return Redirect("~/Home/Index");
                        }

                    }
                    else
                    {
                        if (Url.IsLocalUrl(returnUrl) && returnUrl.Length > 1 && returnUrl.StartsWith("/")
                            && !returnUrl.StartsWith("//") && !returnUrl.StartsWith("/\\"))
                        {
                            return Redirect(returnUrl);
                        }
                        else
                        {
                            if (us.FirstOrDefault(x => x.Id == us.First().Id).Quyen.Equals("NV"))
                            {
                                FormsAuthentication.SetAuthCookie("nhanvien", true);
                            }
                            else
                            {
                                FormsAuthentication.SetAuthCookie(acc.Taikhoan, true);
                            }
                            Session["Taikhoan"] = us.FirstOrDefault(x => x.Id == us.First().Id);
                            return Redirect("~/Home/Index");
                        }
                    }
                }
                else
                {
                    ModelState.AddModelError("", "Tài khoản hoặc mật khẩu không đúng");
                }
            }
            return View("Index");
        }
        public ActionResult Dangky()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Dangky(account acc)
        {
            if (ModelState.IsValid)
            {
                var checkemail = dbcontext.accounts.FirstOrDefault(x => x.Email == acc.Email);
                var checksdt = dbcontext.accounts.FirstOrDefault(x => x.Sdt == acc.Sdt);
                var user = dbcontext.accounts.FirstOrDefault(x => x.Taikhoan == acc.Taikhoan);

                if (user != null)
                {
                    ModelState.AddModelError("", "Tài khoản đã tồn tại");
                }
                else
                {
                    if (checkemail != null || checksdt != null)
                    {
                        if (Request.Form.Count > 0)
                        {
                            String email = Request.Form["Email"];
                            String sdt = Request.Form["Sdt"];
                            String taikhoan = Request.Form["Taikhoan"];
                            String matkhau = Request.Form["Matkhau"];
                            String ten = Request.Form["Ten"];
                            housecareDBContextDataContext context = new housecareDBContextDataContext();
                            account account = new account();
                            if (checkemail.Email.Equals(email) && checksdt.Sdt.Equals(sdt))
                            {
                                account.Id = checksdt.Id;
                                account.Taikhoan = taikhoan;
                                account.Email = email;
                                account.Matkhau = matkhau;
                                account.Ten = ten;
                                account.Sdt = sdt;
                                account.Quyen = "ND";
                                Session["xacnhantaikhoan"] = account;
                                return RedirectToAction("xacnhandangky");
                            }
                            else
                            {
                                if (checkemail.Email.Equals(email))
                                {
                                    account.Id = checkemail.Id;
                                    account.Taikhoan = taikhoan;
                                    account.Email = email;
                                    account.Matkhau = matkhau;
                                    account.Ten = ten;
                                    account.Sdt = checksdt.Sdt;
                                    account.Quyen = "ND";
                                    ViewBag.tk = account;
                                    return RedirectToAction("xacnhandangky");
                                }
                                else
                                {
                                    account.Id = checksdt.Id;
                                    account.Taikhoan = taikhoan;
                                    account.Email = checkemail.Email;
                                    account.Matkhau = matkhau;
                                    account.Ten = ten;
                                    account.Sdt = sdt;
                                    account.Quyen = "ND";
                                    ViewBag.tk = account;
                                    return RedirectToAction("xacnhandangky");

                                }
                            }
                        }
                    }
                    else
                    {
                        if (Request.Form.Count > 0)
                        {
                            String taikhoan = Request.Form["Taikhoan"];
                            String email = Request.Form["Email"];
                            String matkhau = Request.Form["Matkhau"];
                            String ten = Request.Form["Ten"];
                            String sdt = Request.Form["Sdt"];
                            housecareDBContextDataContext context = new housecareDBContextDataContext();
                            account account = new account();
                            account.Taikhoan = taikhoan;
                            account.Email = email;
                            account.Matkhau = matkhau;
                            account.Ten = ten;
                            account.Sdt = sdt;
                            account.Quyen = "ND";
                            context.accounts.InsertOnSubmit(account);
                            context.SubmitChanges();
                            Session["dangky"] = "Đăng ký thành công";
                            return RedirectToAction("Dangky");

                        }
                    }
                }
            }
            return View();
        }
        public ActionResult xacnhandangky()
        {
            if (Request.Form.Count > 0)
            {
                int id = int.Parse(Request.Form["Id"]);
                String taikhoan = Request.Form["Taikhoan"];
                String email = Request.Form["Email"];
                String matkhau = Request.Form["Matkhau"];
                String ten = Request.Form["Ten"];
                String sdt = Request.Form["Sdt"];
                housecareDBContextDataContext context = new housecareDBContextDataContext();
                account account = context.accounts.FirstOrDefault(x => x.Id == id);
                account.Taikhoan = taikhoan;
                account.Mataikhoan = "user080808";
                account.Email = email;
                account.Matkhau = matkhau;
                account.Ten = ten;
                account.Sdt = sdt;
                account.Quyen = "ND";
                context.SubmitChanges();
                return RedirectToAction("Index");
            }
            return View();
        }
        public ActionResult Logout()
        {
            FormsAuthentication.SignOut();
            Session.Clear();
            return View("Index");
        }
    }
}