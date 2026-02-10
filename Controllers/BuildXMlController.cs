using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Xml.Linq;
using Taxweb.Models;
namespace Taxweb.Controllers
{
   
    public class BuildXMlController : Controller
    {
       

        // GET: BuildXMl
        public ActionResult Index()
        {
            return View();
        }
        private const string TAX_NS = "http://kekhaithue.gdt.gov.vn/TKhaiThue";

        private List<XmlEditVM> BuildEditModel(XDocument xdoc)
        {
            XNamespace ns = TAX_NS;

            var result = new List<XmlEditVM>();

            var parent = xdoc
                .Descendants(ns + "CTieuTKhaiChinh")
                .FirstOrDefault();

            if (parent == null)
                return result;

            foreach (var el in parent.Elements())
            {
                result.Add(new XmlEditVM
                {
                    Label = $"Chỉ tiêu [{el.Name.LocalName.Replace("ct", "")}]",
                    XPath = $"/ns:HSoThueDTu/ns:HSoKhaiThue/ns:CTieuTKhaiChinh/ns:{el.Name.LocalName}",
                    Value = el.Value
                });
            }

            return result;
        }


        [HttpPost]
        public ActionResult UploadXml(HttpPostedFileBase xmlFile)
        {
            if (xmlFile == null || xmlFile.ContentLength == 0)
                return View("UploadXml");

            XDocument xdoc;
            using (var stream = xmlFile.InputStream)
            {
                xdoc = XDocument.Load(stream);
            }

            // Lưu XML vào Session (tạm thời)
            Session["XML_DOC"] = xdoc;

            var model = BuildEditModel(xdoc);

            return View("EditXml", model);
        }
    }
}