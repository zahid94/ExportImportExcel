using ExportImportExcel.Models;
using Ganss.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ExportImportExcel.Controllers
{
    public class ExportImportController : Controller
    {
        // GET: ExportImport
        public ActionResult Index()
        {
            return View();
        }

        [HttpGet]
        public ActionResult Export()
        {
            return View();
        }

        [HttpGet]
        public ActionResult ExportExcel()
        {
            var student = new List<Student>
            {
                new Student{Name="zahid",Address="mirpur"},
                new Student { Name="akram",Address="abdullahpur"}
            };
            ExcelMapper mapper = new ExcelMapper();
            var CreateFilePath = @"D:\All Problems\newFile.xlsx";
            if (CreateFilePath !=null)
            {

            }
            //for this mapper install nuget mapper
            mapper.Save(CreateFilePath, student, "SheetName", true);
            TempData["sm"] = "Successfully excel create.";
            return RedirectToAction("Export");
        }
    }    
}