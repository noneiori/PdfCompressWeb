using PdfCompressWeb.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace PdfCompressWeb.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            List<SelectListItem> selectListItems = new List<SelectListItem>();
            selectListItems.Add(new SelectListItem { Text="jpeg" });
            selectListItems.Add(new SelectListItem { Text="png" });
            ViewData["selectListItems"] = selectListItems;
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        [HttpPost]
        public ActionResult UploadFile(HttpPostedFileBase file,string imageType,bool resize)
        {
            var fs = Request.Files;

            if (fs.Count > 1)
            {
                ViewBag.Error = "僅能上傳一個PDF檔案";
                return View();
            }
            if (fs.Count == 0)
            {
                ViewBag.Error = "請上傳一個PDF檔案";
                return View();
            }            

            var handlePdf = fs[0];

            if (handlePdf.ContentType != "application/pdf")
            {
                ViewBag.Error = "僅能上傳PDF檔案";
                return View("Index");
            }

            CompressPdf compress = new CompressPdf();
            string guidname = Guid.NewGuid() + ".pdf";
            string filename = compress.UploadFolder + guidname;
            handlePdf.SaveAs(filename);

            compress.ImageType = imageType;
            compress.ResizePicture = resize;
            string outputFile = compress.CompressFile(guidname);

            return File(compress.UploadFolder + outputFile, "application/pdf");
        }
    }
}