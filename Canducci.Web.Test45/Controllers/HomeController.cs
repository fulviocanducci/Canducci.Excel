using Canducci.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace Canducci.Web.Test45.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
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

        [Route("excel")]
        public FileContentResult Excel()
        {
            var Items = new[]
            {
                new {Id  = 1, Name = "Id 1", Value = 150.00M},
                new {Id  = 2, Name = "Id 2", Value = 250.25M},
                new {Id  = 3, Name = "Id 3", Value = 450.00M},
                new {Id  = 4, Name = "Id 4", Value = 600.00M},
                new {Id  = 5, Name = "Id 5", Value = 700.25M},
                new {Id  = 6, Name = "Id 6", Value = 800.00M}
            }
            .ToArray();

            byte[] ArrayOfBytes = Items.ToExcelByte();

            return File(ArrayOfBytes, ContentTypeExcel.Xlsx.ToValue(), $"{System.Guid.NewGuid()}.xlsx");
        }
    }
}