using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Canducci.Web.TestNETStandard2._0.Models;
using Canducci.Excel;

namespace Canducci.Web.TestNETStandard2._0.Controllers
{
    public class HomeController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult About()
        {
            ViewData["Message"] = "Your application description page.";

            return View();
        }

        public IActionResult Contact()
        {
            ViewData["Message"] = "Your contact page.";

            return View();
        }

        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        [Route("excel")]
        public FileContentResult Excel()
        {
            var Items = new[]
            {
                new {Id  = 1, Name = "Id 1", Value = 150.00M, Data = DateTime.Now, Status = true},
                new {Id  = 2, Name = "Id 2", Value = 250.25M, Data = DateTime.Now, Status = false},
                new {Id  = 3, Name = "Id 3", Value = 450.00M, Data = DateTime.Now, Status = false},
                new {Id  = 4, Name = "Id 4", Value = 600.00M, Data = DateTime.Now, Status = false},
                new {Id  = 5, Name = "Id 5", Value = 1700.25M, Data = DateTime.Now, Status = false},
                new {Id  = 6, Name = "Id 6", Value = 800.00M, Data = DateTime.Now, Status = true},
                new {Id  = 7, Name = "Id 7", Value = 150.00M, Data = DateTime.Now, Status = false},
                new {Id  = 8, Name = "Id 8", Value = 250.25M, Data = DateTime.Now, Status = true},
                new {Id  = 9, Name = "Id 9", Value = 1450.00M, Data = DateTime.Now, Status = false},
                new {Id  = 10, Name = "Id 10", Value = 1600.00M, Data = DateTime.Now, Status = false},
                new {Id  = 11, Name = "Id 11", Value = 1700.25M, Data = DateTime.Now, Status = true},
                new {Id  = 12, Name = "Id 12", Value = 1800.00M, Data = DateTime.Now, Status = false}
            }
            .ToArray();            
            byte[] ArrayOfBytes = Items.ToExcelByte(x =>
            {
                x.DateFormat = "dd/mm/yyyy hh:mm";
                x.Headers = HeaderCreator
                    .Create()
                    .Add("Número", 1).Add("Nome Completo", 2)
                    .Add("Valor", 3).Add("Data Atual", 4)
                    .Add("Ativo", 5)
                    .Render();
                    
            });

            return File(ArrayOfBytes, ContentTypeExcel.Xlsx.ToValue(), $"{Guid.NewGuid()}.xlsx");
        }
    }
}
