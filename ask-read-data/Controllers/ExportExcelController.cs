using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ask_read_data.Models;
using ask_read_data.Repository;
using Microsoft.AspNetCore.Mvc;

namespace ask_read_data.Controllers
{
    public class ExportExcelController : Controller
    {
        private readonly IExportExcel _exportExcel;
        public ExportExcelController(IExportExcel exportExcel)
        {
            this._exportExcel = exportExcel;
        }
        [HttpGet]
        public IActionResult ExportExcel()
        {
            var model = new ExportExcelViewModel() { Yotebi = DateTime.Today, DataTableBody = _exportExcel.GetAll2000DataImport() };
            return View(model);
        }

        [HttpPost]
        public IActionResult ExportExcel(ExportExcelViewModel model)
        {
            return View(model);
        }
    }
}
