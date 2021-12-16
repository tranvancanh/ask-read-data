using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ask_read_data.Models.ViewModel;
using ask_read_data.Repository;
using mclogi.common;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace ask_read_data.Controllers
{
    [Authorize]
    public class ExportExcelController : Controller
    {
        public const string FLOOR_ASSY = "FLOOR ASSY";
        public const string FLAME_ASSY = "FLAME ASSY";
        public const int PALETNO_FLOOR_ASSY = 8;
        public const int PALETNO_FLAME_ASSY = 4;
        public const string ASHITA_IKO = "明日以降";

        private readonly IExportExcel _excelExport;
        public ExportExcelController(IExportExcel excelExport)
        {
            this._excelExport = excelExport;
        }
        [HttpGet]
        public IActionResult ExportExcel()
        {
            return View(new ExportExcelViewModel());
        }

        [HttpPost]
        public IActionResult ExportExcel(ExportExcelViewModel modelRequset, string floorassy, string flameassy)
        {
            if (!ModelState.IsValid)
            {
                return RedirectToAction("ExcelExport", "ExcelExport");
            }
            var Claims = User.Claims.ToList();
            var dt1 = new DataTable();
            var dt2 = new DataTable();
            var filename = string.Empty;
            var BubanMeiType = string.Empty;
            List<string> sheetName = new List<string>();
            /////////////////////////////////////////////////////////////////////////////////////////////////////
            string button = floorassy + flameassy;
            try
            {
                switch (button)
                {
                    case "floorassy":
                        {
                            BubanMeiType = FLOOR_ASSY;
                            var dt = _excelExport.GetFloor_Flame_Assy(modelRequset.Floor_Assy, FLOOR_ASSY);
                            dt1 = dt.Item1;
                            dt2 = dt.Item2;
                            filename = "FloorAssy_";
                            sheetName = new List<string>() { "FloorAssy_⓵Sheet1", "FloorAssy_⓶Sheet2" };
                            break;
                        }
                    case "flameassy":
                        {
                            BubanMeiType = FLAME_ASSY;
                            var dt = _excelExport.GetFloor_Flame_Assy(modelRequset.Flame_Assy, FLAME_ASSY);
                            dt1 = dt.Item1;
                            dt2 = dt.Item2;
                            filename = "Flame_Assy_";
                            sheetName = new List<string>() { "Flame_Assy_⓵Sheet1", "Flame_Assy_⓶Sheet2" };
                            break;
                        }
                    default:
                        {
                            //dt = null;
                            ViewData["error"] = "ボタンがおかしいです";
                            return View();
                        }
                }
            }
            catch(Exception ex)
            {
                var error = ex.Message;
                //var GetType = ex.GetType;
                TempData["error"] = "ダウンロードに失敗しました " + error + "| Exception Type: " + ex.GetType().ToString();
                return View();
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
                if (util.ExportExcel(dt1, dt2, expPath, null, null,sheetName))
                {

                    // ダウンロード履歴登録
                    var outputFilename = string.Concat(filename, Utility.EXTENSION_XLSX);
                    // InsertFile_Update_Log(UPDOWNKUBUN_DWN, outputFilename);

                    // 出力に成功したら、ファイルダウンロード
                    var file = System.IO.File.ReadAllBytes(expPath);
                    if(_excelExport.RecordDownloadHistory(ref dt1, BubanMeiType, Claims) > 0)
                    {
                        TempData["success"] = "ダウンロードに成功しました";
                        return File(file, System.Net.Mime.MediaTypeNames.Application.Octet, outputFilename);
                    }
                    TempData["error"] = "ダウンロードに失敗しました";
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
                //var GetType = ex.GetType;
                TempData["error"] = "ダウンロードに失敗しました " + error + "| Exception Type: " + ex.GetType().ToString();
                return View();
            }
            
           // return View(modelRequset);
        }
    }
}
