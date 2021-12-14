using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ask_read_data.Models.ViewModel;
using ask_read_data.Repository;
using mclogi.common;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace ask_read_data.Controllers
{
    public class ExcelExportController : Controller
    {
        private readonly IExcelExport _excelExport;
        public ExcelExportController(IExcelExport excelExport)
        {
            this._excelExport = excelExport;
        }
        [HttpGet]
        public IActionResult ExcelExport()
        {
            return View(new ExcelExportViewModel());
        }

        [HttpPost]
        public IActionResult ExcelExport(ExcelExportViewModel modelRequset, string floorassy, string flameassy)
        {
            if (!ModelState.IsValid)
            {
                return RedirectToAction("ExcelExport", "ExcelExport");
            }
            DataTable dt = null;
            var filename = "";
            /////////////////////////////////////////////////////////////////////////////////////////////////////
            string button = floorassy + flameassy;
            switch (button)
            {
                case "floorassy":
                    {
                        dt = _excelExport.GetFloor_Flame_Assy(modelRequset.Floor_Assy, "FLOOR ASSY");
                        filename = "FloorAssy_";
                        break;
                    }
                case "flameassy":
                    {
                        dt = _excelExport.GetFloor_Flame_Assy(modelRequset.Flame_Assy, "FLAME ASSY");
                        filename = "Flame_Assy";
                        break;
                    }
                default:
                    {
                        dt = null;
                        ViewData["error"] = "ボタンがおかしいです";
                        return View();
                    }
            }
            // 出力フォルダパス
            var rootPath = Directory.GetCurrentDirectory();
            var path = Path.Combine(rootPath, @"wwwroot\DownloadFiles");

            /////////////////////////////////////////////////////////////////////////////////////////////////////
            // ファイル名
            var tmpFilename = string.Concat(filename,
                "_",
                HttpContext.Session.GetString(Utility.SESSION_KEY_USERCD),
                "_",
                DateTime.Now.ToString("yyyyMMddHHmmssfff"),
                "_",
                Path.ChangeExtension(Path.GetRandomFileName(), Utility.EXTENSION_XLSX)
                );

            // 出力パス
            var expPath = Path.Combine(path, tmpFilename);

            try
            {
                // Excelファイル生成
                Utility util = new Utility();
                if (util.ExportExcel(dt, expPath, null, null))
                {

                    // ダウンロード履歴登録
                    var outputFilename = string.Concat(filename, Utility.EXTENSION_XLSX);
                    // InsertFile_Update_Log(UPDOWNKUBUN_DWN, outputFilename);

                    // 出力に成功したら、ファイルダウンロード
                    var file = System.IO.File.ReadAllBytes(expPath);
                    TempData["success"] = "ダウンロードに成功しました";
                    return File(file, System.Net.Mime.MediaTypeNames.Application.Octet, outputFilename);
                }
                else
                {
                    TempData["error"] = "ダウンロードに失敗しました";
                    return View();
                }
            }
            catch(Exception ex)
            {
                var error = ex.Message;
                TempData["error"] = "ダウンロードに失敗しました";
                return View();
            }
            
           // return View(modelRequset);
        }
    }
}
