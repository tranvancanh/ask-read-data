using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ask_read_data.Models.ViewModel;
using ask_read_data.Repository;
using mclogi.common;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;

namespace ask_read_data.Controllers
{
    [Authorize]
    public class ExportExcelController : Controller
    {
        /****************************部品番号*************************************/
        public const string BUHIN_FLOOR_74300WL20P = "74300WL20P";
        public const string BUHIN_FLOOR_74300WL30P = "74300WL30P";

        public const string BUHIN_FLAME_743B2W000P = "743B2W000P";
        public const string BUHIN_FLAME_743B2W010P = "743B2W010P";

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

        public const int STEP_PAGE = 38; 

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
            var tempFile = "";
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
                            tempFile = @"wwwroot\FormatFile\ASUKA_FL00R_ASSY_FORMAT_序列表.xlsx";
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
                            tempFile = @"wwwroot\FormatFile\ASUKA_FRAME_ASSY_FORMAT_序列表.xlsx";
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
            string tempPath = Path.Combine(rootPath, tempFile);
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
                if (util.ExportExcel(dt1, dt2, expPath, tempPath, BubanMeiType, null, null, sheetName))
                {
                    //if (true) 
                    //{
                    //var newFile = new FileInfo(expPath);
                    //var tempFile = new FileInfo(tempPath);
                    ///////////////////////////////////////////////////////////////////////////////////////////////////////
                    //using (var package = new ExcelPackage(newFile, tempFile))
                    //{
                    //    //ExcelWorksheet worksheet = null;
                    //    //acquisition of WorkSheet object. Because it FirstOrDefault (), it is better that originally was null check at a later stage. 
                    //    var worksheet = package.Workbook.Worksheets.Where(s => s.Name == sheetName[0]).FirstOrDefault();
                    //    //first dimension when specifying a cell in a two dimensional array Y direction, the second dimension is the X direction. 

                    //    worksheet.Cells[1, 2].Value = "template of sample";
                    //    worksheet.Cells[1, 2].Style.Font.Color.SetColor(Color.Red);
                    //    for (int i = 0; i < 14; i++)
                    //    {
                    //        worksheet.Cells[4 + i, 1].Value = DateTime.Now.Hour;
                    //        worksheet.Cells[4 + i, 2].Value = DateTime.Now.Minute;
                    //        worksheet.Cells[4 + i, 3].Value = DateTime.Now.Second;
                    //        worksheet.Cells[4 + i, 4].Value = DateTime.Now.Millisecond;
                    //        if((4 + i)%4 == 0)
                    //        {
                    //            worksheet.Row(4 + i).PageBreak = true;
                    //            // Insert a page break after the tenth row.
                    //            // Insert a page break below the 14th row.
                    //            //worksheet.HorizontalPageBreaks.Add(14);

                    //            //workbook.Worksheets[0].HPageBreaks.Add(worksheet.Range["A7"]);
                    //            //workbook.Worksheets[0].HPageBreaks.Add(worksheet.Range["A13"]);
                    //            //workbook.Worksheets[0].ViewMode = ViewMode.Preview;
                    //        }
                    //    }
                    //    worksheet.PrinterSettings.PrintArea = worksheet.Cells["A:1,A:15"];
                    //    //ExcelWorksheet.PrinterSettings.PrintArea = ExcelWorksheet.Cells[1, 1, 20, 4];
                    //    //save the Excel file 
                    //    package.Save();
                    //}

                    //////////////////////////////////////////////////////////////////////////////////////////////////////
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
