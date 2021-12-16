﻿using System;
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
        /****************************部番略式記号*************************************/
        public const string FLOOR_ASSY = "FL00R ASSY";
        public const string FLAME_ASSY = "FRAME ASSY";

        /****************************入数*************************************/
        public const int PALETNO_FLOOR_ASSY = 8;
        public const int PALETNO_FLAME_ASSY = 4;

        public const string ASHITA_IKO = "明日以降";
        /****************************シート名***********************************/
        /****FLOOR ASSY****/
        public const string FLOORASSY_SHEET1 = "フロアー出荷用";
        public const string FLOORASSY_SHEET2 = "フロアー生産用";
        /****FRAME ASSY****/
        public const string FRAMEASSY_SHEET1 = "フレーム出荷用";
        public const string FRAMEASSY_SHEET2 = "フレーム生産用";

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
                            filename = "FLOOR_ASSY_" + modelRequset.Floor_Assy.Year.ToString().Substring(2,2) + modelRequset.Floor_Assy.Month.ToString() + modelRequset.Floor_Assy.Day.ToString() + "出荷分序列データー";
                            sheetName = new List<string>() { FLOORASSY_SHEET1, FLOORASSY_SHEET2 };
                            break;
                        }
                    case "flameassy":
                        {
                            BubanMeiType = FLAME_ASSY;
                            var dt = _excelExport.GetFloor_Flame_Assy(modelRequset.Flame_Assy, FLAME_ASSY);
                            dt1 = dt.Item1;
                            dt2 = dt.Item2;
                            filename = "FLAME_ASSY_" + modelRequset.Flame_Assy.Year.ToString().Substring(2, 2) + modelRequset.Flame_Assy.Month.ToString() + modelRequset.Flame_Assy.Day.ToString() + "出荷分序列データー";
                            sheetName = new List<string>() { FRAMEASSY_SHEET1, FRAMEASSY_SHEET2 };
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
                if(dt1.Rows.Count <= 0 || dt2.Rows.Count <=0)
                {

                    TempData["error"] = "その日付はデータが存在していませんのでダウンロードに失敗しました";
                    return View();
                }
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
                        TempData["error"] = null;
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
