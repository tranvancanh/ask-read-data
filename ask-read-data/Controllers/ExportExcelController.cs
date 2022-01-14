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
        public IActionResult ExportExcel(ExportExcelViewModel viewModel)
        {
            //var viewModel = new ExportExcelViewModel();
            var bubantype = "";
            bubantype = FL00R_ASSY;
            var result1 = _excelExport.FindPositionParetoRenbanLasttime(bubantype);
            //var result1 = _excelExport.FindPositionParetoRenban(viewModel.Floor_Assy, bubantype);
            viewModel.Floor_Assy = result1.Item1;
            viewModel.Floor_Position = result1.Item2;
            viewModel.Floor_ParetoRenban = result1.Item3;
            bubantype = FRAME_ASSY;
            var result2 = _excelExport.FindPositionParetoRenbanLasttime(bubantype);
            //var result2 = _excelExport.FindPositionParetoRenban(viewModel.Flame_Assy, bubantype);
            viewModel.Flame_Assy = result2.Item1;
            viewModel.Flame_Position = result2.Item2;
            viewModel.Flame_ParetoRenban = result2.Item3;

            //viewModel.SelectList = new Microsoft.AspNetCore.Mvc.Rendering.SelectList(list);
            viewModel.ListData = new DownloadHistoryViewModel() { DataTableHeader = new Models.Entity.FileDownloadLogModel(), DataTableBody = _excelExport.FindDownloadHistory(viewModel.SearchDate) };
             
            return View(viewModel);
        }

        [HttpPost]
        public IActionResult ExportExcel(DateTime Floor_Assy = new DateTime(),  int Floor_Position = 0, int Floor_ParetoRenban = 0, DateTime Flame_Assy = new DateTime(), int Flame_Position = 0, int Flame_ParetoRenban = 0, DateTime SearchDate = new DateTime(), string clickbtn = "")
        {
            var modelRequset = new ExportExcelViewModel() { 
                                                             Floor_Assy = Floor_Assy ,
                                                             Floor_Position = Floor_Position,
                                                             Floor_ParetoRenban = Floor_ParetoRenban,
                                                             Flame_Assy = Flame_Assy,
                                                             Flame_Position = Flame_Position,
                                                             Flame_ParetoRenban = Flame_ParetoRenban,
                                                             SearchDate = SearchDate
                                                          };
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
            try
            {
                switch (clickbtn)
                {
                    case "floorassy":
                        {
                            BubanMeiType = FL00R_ASSY;
                            var dt = _excelExport.GetFloor_Flame_Assy(modelRequset, FL00R_ASSY);
                            dt1 = dt.Item1;
                            dt2 = dt.Item2;
                            filename = "FLOOR_ASSY_出荷分序列データー_" + modelRequset.Floor_Assy.Year.ToString() + modelRequset.Floor_Assy.ToString("MM") + modelRequset.Floor_Assy.ToString("dd") + "(" + modelRequset.Floor_Position.ToString() + ")(" + modelRequset.Floor_ParetoRenban.ToString() + ")";
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
                            filename = "FLAME_ASSY_出荷分序列データー_" + modelRequset.Flame_Assy.Year.ToString() + modelRequset.Flame_Assy.ToString("MM") + modelRequset.Flame_Assy.ToString("dd") + "(" + modelRequset.Flame_Position.ToString() + ")(" + modelRequset.Flame_ParetoRenban.ToString() + ")";
                            sheetName = new List<string>() { FRAMEASSY_SHEET1, FRAMEASSY_SHEET2 };
                            tempFile = ASUKA_FRAME_ASSY_TEMPLATE;
                            break;
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
                return View(ReturnDataViewModel(modelRequset));
            }
            // 出力フォルダパス
            var rootPath = Directory.GetCurrentDirectory();
            var path = Path.Combine(rootPath, @"wwwroot\DownloadFiles");
            //Webサーバーのdownloadフォルダーがない場合は作成
            if (!System.IO.Directory.Exists(path))
            {
                System.IO.Directory.CreateDirectory(path);
            }
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
                    return Json(new { StatusCode = false, Mess = TempData["error"] });
                }
                // Excelファイル生成
                Utility util = new Utility();
                if (util.ExportExcel(dt1, dt2, expPath, tempPath, BubanMeiType, null, null, sheetName))
                {
                    //////////////////////////////////////////////////////////////////////////////////////////////////////
                    var outputFilename = string.Concat(filename, Utility.EXTENSION_XLSX);
                    // InsertFile_Update_Log(UPDOWNKUBUN_DWN, outputFilename);
                    // ダウンロード履歴登録
                    // 出力に成功したら、ファイルダウンロード
                    var file = System.IO.File.ReadAllBytes(expPath);
                    if (_excelExport.RecordDownloadHistory(ref dt1, BubanMeiType, Claims) > 0)
                    {
                        TempData["error"] = null;
                        TempData["success"] = "ダウンロードに成功しました";
                        //return File(file, System.Net.Mime.MediaTypeNames.Application.Octet, outputFilename, true);
                        return Json(new { StatusCode = true, fileName = tmpFilename, Mess = TempData["success"] });
                    }
                    TempData["error"] = "ダウンロードに失敗しました";
                    return Json(new { StatusCode = false, Mess = TempData["error"] });
                }
                else
                {
                    TempData["error"] = "ダウンロードに失敗しました";
                    return Json(new { StatusCode = false, Mess = TempData["error"] });
                }
            }
            catch (Exception ex)
            {
                var error = ex.Message;
                //var GetType = ex.GetType;
                TempData["error"] = "ダウンロードに失敗しました " + error + "| Exception Type: " + ex.GetType().ToString();
                return Json(new { StatusCode = false, Mess = TempData["error"] });
            }

            // return View(modelRequset);
        }

        [HttpPost]
        public ActionResult Search(ExportExcelViewModel modelRequset)
        {
            return View("ExportExcel", ReturnDataViewModel(modelRequset));
        }

        [HttpGet]
        public ActionResult Download(string fileName)
        {
            //Get the temp folder and file path in server
            // 出力フォルダパス
            var rootPath = Directory.GetCurrentDirectory();
            var path = Path.Combine(rootPath, @"wwwroot\DownloadFiles");

            string fullPath = Path.Combine(path, fileName);
            byte[] fileByteArray = System.IO.File.ReadAllBytes(fullPath);
            //System.IO.File.Delete(fullPath);
            return File(fileByteArray, "application/vnd.ms-excel", fileName);
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
        private ExportExcelViewModel ReturnDataViewModel(ExportExcelViewModel viewModel)
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
