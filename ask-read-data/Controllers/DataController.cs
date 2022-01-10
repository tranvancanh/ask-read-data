using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ask_read_data.Models;
using ask_read_data.Repository;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace ask_read_data.Controllers
{
    [Authorize]
    public class DataController : Controller
    {
        private readonly IData _exportExcel;
        public DataController(IData exportExcel)
        {
            this._exportExcel = exportExcel;
        }
        [HttpGet]
        public IActionResult Data()
        {
            var model = new DataViewModel1() { ImportDate = DateTime.Today, DataTableBody = _exportExcel.SearchDataImport(DateTime.Today) };
            return View(model);
        }

        [HttpPost]
        public IActionResult Data(DataViewModel1 model)
        {
            var resultSearch = new DataViewModel1() { ImportDate = model.ImportDate, DataTableBody = _exportExcel.SearchDataImport(model.ImportDate) };
            return View(resultSearch);
        }
    }
}
