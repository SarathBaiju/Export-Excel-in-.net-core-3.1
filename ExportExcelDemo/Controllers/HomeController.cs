using ClosedXML.Excel;
using ExportExcelDemo.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ExportExcelDemo.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            var users = GetDummyUsers();
            return View(users);
        }

       public async Task<IActionResult> ExportExcel()
        {
            try
            {
                using (var workBook = new XLWorkbook())
                {
                    var workSheet = workBook.Worksheets.Add("users");
                    var currentRow = 1;
                    workSheet.Cell(currentRow, 1).Value = "Id";
                    workSheet.Cell(currentRow, 2).Value = "Name";

                    var users = GetDummyUsers();
                    foreach (var user in users)
                    {
                        currentRow++;
                        workSheet.Cell(currentRow, 1).Value = user.Id;
                        workSheet.Cell(currentRow, 2).Value = user.Name;
                    }

                    using (var stream = new MemoryStream())
                    {
                        workBook.SaveAs(stream);
                        var content = stream.ToArray();
                        return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "user.xlsx");
                    }
                }

                return Ok();
            }
            catch (Exception)
            {

                throw;
            }
        }
        private List<User> GetDummyUsers()
        {
            return new List<User>
            {
                new User{Id=1, Name= "user1"},
                new User{Id=2, Name= "user2"},
                new User{Id=3, Name= "user3"}
            };
        }
    }

    public class User
    {
        public int Id { get; set; }
        public string Name { get; set; }
    }
}
