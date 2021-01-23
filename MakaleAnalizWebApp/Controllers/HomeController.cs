using MakaleAnalizWebApp.Models;
using MakaleAnalizWebApp.Service;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace MakaleAnalizWebApp.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }
        [HttpPost]
        public ActionResult Analiysis(HttpPostedFileBase file)
        {
            try
            {
                string path = Path.Combine(Server.MapPath("~/files"),
                                Path.GetFileName(DateTime.Now.ToString("yyyy-MM-dd-ss-mm") +file.FileName));
                file.SaveAs(path);
                Analiysis analiysis = new Analiysis(path);
                analiysis.checkFile();

                return View(analiysis.results);
            }
            catch (Exception)
            {

                return View(new List<Result>()
                { new Result()
                {
                    message="Okuma İşlemi Sırasında Bir hata oluştu lütfen tekrar deneyin.",
                    isSuccess=false }});
            }

}

    }
}