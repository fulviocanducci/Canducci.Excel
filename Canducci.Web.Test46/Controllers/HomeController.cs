using System.Linq;
using System.Web.Mvc;
using Canducci.Excel;
namespace Canducci.Web.Test46.Controllers
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
                new {Id  = 3, Name = "Id 3", Value = 450.00M}
            }
            .ToArray();

            byte[] ArrayOfBytes = Items.ToExcelByte();

            return File(ArrayOfBytes, ContentTypeExcel.Xlsx.ToValue(), $"{System.Guid.NewGuid()}.xlsx");
        }
    }
}