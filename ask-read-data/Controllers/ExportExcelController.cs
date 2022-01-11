using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ask_read_data.Commons;
using ask_read_data.Models.ViewModel;
using ask_read_data.Repository;
using mclogi.common;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Diagnostics;
using Newtonsoft.Json;
using Microsoft.AspNetCore.Mvc.Rendering;

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
        public const string FL00R_ASSY = "FL00R ASSY";
        public const string FRAME_ASSY = "FRAME ASSY";

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

        /****************************Template file***********************************/
        private const string ASUKA_FL00R_ASSY_TEMPLATE = @"wwwroot\FormatFile\ASUKA_FL00R_ASSY_TEMPLATE_序列表.xlsx";
        private const string ASUKA_FRAME_ASSY_TEMPLATE = @"wwwroot\FormatFile\ASUKA_FRAME_ASSY_TEMPLATE_序列表.xlsx";

        public const int STEP_PAGE = 38;

        private readonly IExportExcel _excelExport;
        public ExportExcelController(IExportExcel excelExport)
        {
            this._excelExport = excelExport;
        }
        [HttpGet]
        public IActionResult ExportExcel()
        {
            var viewModel = new ExportExcelViewModel();
            var bubantype = "";
            bubantype = FL00R_ASSY;
            var result1 = _excelExport.FindPositionParetoRenban(viewModel.Floor_Assy, bubantype);
            viewModel.Floor_Position = result1.Item1;
            viewModel.Floor_ParetoRenban = result1.Item2;
            bubantype = FRAME_ASSY;
            var result2 = _excelExport.FindPositionParetoRenban(viewModel.Flame_Assy, bubantype);
            viewModel.Flame_Position = result2.Item1;
            viewModel.Flame_ParetoRenban = result2.Item2;

            //viewModel.SelectList = new Microsoft.AspNetCore.Mvc.Rendering.SelectList(list);
            viewModel.ListData = new DownloadHistoryViewModel() { DataTableHeader = new Models.Entity.FileDownloadLogModel(), DataTableBody = _excelExport.FindDownloadHistory(viewModel.SearchDate) };
             
            return View(viewModel);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public IActionResult ExportExcel(ExportExcelViewModel modelRequset, string floorassy = null, string flameassy = null, string searchbtn = null )
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
            string button = floorassy + flameassy + searchbtn;
            try
            {
                switch (button)
                {
                    case "floorassy":
                        {
                            BubanMeiType = FL00R_ASSY;
                            var dt = _excelExport.GetFloor_Flame_Assy(modelRequset, FL00R_ASSY);
                            dt1 = dt.Item1;
                            dt2 = dt.Item2;
                            filename = "FLOOR_ASSY_出荷分序列データー_" + modelRequset.Floor_Assy.Year.ToString().Substring(2, 2) + modelRequset.Floor_Assy.ToString("MM") + modelRequset.Floor_Assy.ToString("dd");
                            sheetName = new List<string>() { FLOORASSY_SHEET1, FLOORASSY_SHEET2 };
                            tempFile = ASUKA_FL00R_ASSY_TEMPLATE;
                            break;
                        }
                    case "flameassy":
                        {
                            BubanMeiType = FRAME_ASSY;
                            var dt = _excelExport.GetFloor_Flame_Assy(modelRequset, FRAME_ASSY);
                            dt1 = dt.Item1;
                            dt2 = dt.Item2;
                            filename = "FLAME_ASSY_出荷分序列データー_" + modelRequset.Flame_Assy.Year.ToString().Substring(2, 2) + modelRequset.Flame_Assy.ToString("MM") + modelRequset.Flame_Assy.ToString("dd");
                            sheetName = new List<string>() { FRAMEASSY_SHEET1, FRAMEASSY_SHEET2 };
                            tempFile = ASUKA_FRAME_ASSY_TEMPLATE;
                            break;
                        }
                    case "search":
                        {
                            return View(CopyExportExcelViewModel(modelRequset));
                        }
                    default:
                        {
                            //dt = null;
                            ViewData["error"] = "ボタンがおかしいです";
                            return View(new ExportExcelViewModel());
                        }
                }
            }
            catch (Exception ex)
            {
                var error = ex.Message;
                //var GetType = ex.GetType;
                TempData["error"] = "ダウンロードに失敗しました " + error + "| Exception Type: " + ex.GetType().ToString();
                return View(new ExportExcelViewModel());
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
                if (dt1.Rows.Count <= 0 || dt2.Rows.Count <= 0)
                {

                    TempData["error"] = "データが存在していませんのでダウンロードに失敗しました";
                    return View(new ExportExcelViewModel());
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
                    if (_excelExport.RecordDownloadHistory(ref dt1, BubanMeiType, Claims) > 0)
                    {
                        TempData["error"] = null;
                        TempData["success"] = "ダウンロードに成功しました";
                        return File(file, System.Net.Mime.MediaTypeNames.Application.Octet, outputFilename, true);
                    }
                    TempData["error"] = "ダウンロードに失敗しました";
                    return File(file, System.Net.Mime.MediaTypeNames.Application.Octet, outputFilename, true);
                }
                else
                {
                    TempData["error"] = "ダウンロードに失敗しました";
                    return View(new ExportExcelViewModel());
                }
            }
            catch (Exception ex)
            {
                var error = ex.Message;
                //var GetType = ex.GetType;
                TempData["error"] = "ダウンロードに失敗しました " + error + "| Exception Type: " + ex.GetType().ToString();
                return View(new ExportExcelViewModel());
            }

            // return View(modelRequset);
        }

        [HttpPost]
        public JsonResult GetPositionParetoRenban(DateTime date, string bubantype)
        {
            var position = 0;
            var renban = 0;
            var result = new Tuple <int,int>( 0, 0);
            try
            {
                switch (bubantype)
                {
                    case "floorassy":
                        {
                            bubantype = FL00R_ASSY;
                            result = _excelExport.FindPositionParetoRenban(date, bubantype);
                            position = result.Item1;
                            renban = result.Item2;
                            break;
                        }
                    case "flameassy":
                        {
                            bubantype = FRAME_ASSY;
                            result = _excelExport.FindPositionParetoRenban(date, bubantype);
                            position = result.Item1;
                            renban = result.Item2;
                            break;
                        }
                    default:
                        {
                            ViewData["error"] = "ボタンがおかしいです";
                            return Json(new { StatusCode = false, Mess = "error" });
                        }
                }
            }
            catch(Exception ex)
            {
                ErrorInfor.DebugWriteLineError(ex);
                Debug.WriteLine("========================================================================================");
                return Json(new { StatusCode = false, Mess = "error" });
            }

            return Json(new { StatusCode = true, position = position, renban = renban });
        }
        private ExportExcelViewModel CopyExportExcelViewModel(ExportExcelViewModel viewModel)
        {
            try
            {
                /*                                 固定分 Start                                              */
                var result1 = _excelExport.FindPositionParetoRenban(viewModel.Floor_Assy, FL00R_ASSY);
                viewModel.Floor_Position = result1.Item1;
                viewModel.Floor_ParetoRenban = result1.Item2;

                var result2 = _excelExport.FindPositionParetoRenban(viewModel.Flame_Assy, FRAME_ASSY);
                viewModel.Flame_Position = result2.Item1;
                viewModel.Flame_ParetoRenban = result2.Item2;
                /*                                 固定分 Stop                                              */

                viewModel.ListData = new DownloadHistoryViewModel() { DataTableHeader = new Models.Entity.FileDownloadLogModel(), DataTableBody = _excelExport.FindDownloadHistory(viewModel.SearchDate) };

            }
            catch(Exception ex)
            {
                var error = ex.Message;
                //var GetType = ex.GetType;
                TempData["error"] = "途中にエラーが発生しました!";
                return new ExportExcelViewModel();
            }
           
            return viewModel;
        }

    }
}
