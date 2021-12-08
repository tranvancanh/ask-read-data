using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace ask_read_data.Controllers
{
    [Authorize]
    public class ImportDataController : Controller
    {
        [HttpGet]
        public IActionResult ImportData()
        {
            return View();
        }

        [HttpPost]
        public IActionResult ImportData(List<IFormFile> FileUpload)
        {
            ViewData.Add("erroremess", null);
            ViewData.Add("successmess", null);
            var files = FileUpload;
            if(files == null || files.Count <= 0)
            {
                ViewData["erroremess"] = "ファイル内にデータが存在しません";
                return View();
            }
            foreach( var file in files)
            {
                if(file.Length > 0)
                {
                    //  ファイル形式のCheck
                    if (file.ContentType != "text/plain")
                    {
                        ViewData["erroremess"] = "ファイル形式が間違っているため、読み取ることができません";
                        return View();
                    }
                }


            }

            ViewData["successmess"] = "ファイル内にデータが保存されました";
            return View();
        }
    }
}
