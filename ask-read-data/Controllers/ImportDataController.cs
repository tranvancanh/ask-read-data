using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using ask_read_data.Areas.Admin.Commons;
using ask_read_data.Commons;
using ask_read_data.Models;
using ask_read_data.Models.Entity;
using ask_read_data.Models.ViewModel;
using ask_read_data.Repository;
using ask_tzn_funamiKD.Commons;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;

namespace ask_read_data.Controllers
{
    [Authorize]
    public class ImportDataController : ReponseMessage
    {
        public static List<string> ListItems = new List<string>();
        /****************************カタカナ名略式記号*************************************/
        private const string KATAKANA_NAME_FLOOR_ASSY = "フロアー";
        private const string KATAKANA_NAME_FLAME_ASSY = "フレーム";
        private const string KATAKANA_NAME_HOKANO = "他の品番名称";

        private readonly IImportData _importData;
        private DataModel dataModel;
        private Bu_MastarModel buMastar;

        private DateTime DateMin = new DateTime(2000, 01, 01, 00, 00, 00);
        private DateTime DateMax = new DateTime(2050, 01, 01, 00, 00, 00);
        public ImportDataController(IImportData importData, DataModel dataModel, Bu_MastarModel buMastar)
        {
            this._importData = importData;
            this.dataModel = dataModel;
            this.buMastar = buMastar;
        }

        [HttpGet]
        public IActionResult ImportData(ImportViewModel viewModel)
        {
            try
            {
                viewModel.ListItem = CheckListItems(viewModel.SearchDate);
                viewModel.ListData = new DataViewModel() { DataTableHeader = new Models.Entity.DataModel(), DataTableBody = _importData.FindDataOfLastTimeInit(DateTime.Today) };
            }
            catch
            {
                return View(new ImportViewModel());
            }
            
            return View(viewModel);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="FileUpload"></param>
        /// <returns></returns>
        [HttpPost]
        public IActionResult ImportData(List<IFormFile> FileUpload)
        {
            ImportViewModel viewModel = new ImportViewModel();
            /////////////////////////////////////////////////////////////////////////////
            //  Get Current User Claims
            //  var abc = new UserInfor().UserInfo().UserName;
            var userName = User.GetLoggedInUserName();
            var userEmail = User.GetLoggedInUserEmail();
            var user = User.Claims.Where(c => c.Type == ClaimTypes.Name).First().Value;
            var Claims = User.Claims.ToList();
            ////////////////////////////////////////////////////////////////////////////
            ViewData.Add("erroremess", null);
            ViewData.Add("successmess", null);
            //  宣言 and レセット
            var result = new List<object>();
            var buMastars = new List<Bu_MastarModel>();
            result.Clear();
            var files = FileUpload;
            var fileNamee = "";
            if (files == null || files.Count <= 0)
            {
                ViewData["erroremess"] = "ファイル内にデータが存在しません";
                return View(ReturnDataView(viewModel));
            }
            try
            {
                foreach (var file in files)
                {
                    fileNamee = file.FileName;
                    if (file.Length > 0)
                    {
                        //  ファイル形式のCheck
                        if (file.ContentType != "text/plain")
                        {
                            ViewData["erroremess"] = "ファイル形式が間違っているため、読み取ることができません";
                            return View(ReturnDataView(viewModel));
                        }

                        //1ファイルデータずつ読み取り
                        ReaderData(file, ref result, ref buMastars);
                        var firstPosition = result.FirstOrDefault().ToString();
                        if (firstPosition == "NG")
                        {
                            var fileName = result[1];
                            var lineNo = result[3];
                            ViewData["successmess"] = null;
                            ViewData["erroremess"] = fileName.ToString() + $@"： ファイル読み込み中にエラーが発生しました
                                            | lineNo: " + lineNo.ToString();
                            return View(ReturnDataView(viewModel));
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                ViewData["successmess"] = null;
                ViewData["erroremess"] = ex.Message + $@" | file name: {fileNamee}";
                return View(ReturnDataView(viewModel));
            }
            ResponResult res = null;
            try
            {
                // go to service
                res = _importData.ImportDataDB(result, Claims, buMastars);
            }
            catch (Exception ex)
            {
                ViewData["successmess"] = null;
                ViewData["erroremess"] = ex.Message + $@" | file name: {fileNamee}";
                return View(ReturnDataView(viewModel));
            }
            if (res.Status != "OK")
            {
                ViewData["successmess"] = null;
                ViewData["erroremess"] = res.Resmess;
                return View(ReturnDataView(viewModel));
            }

            // ViewData["successmess"] = "ファイル内にデータが保存されました";
            ViewData["erroremess"] = null;
            ViewData["successmess"] = res.Resmess;

            return View(ReturnDataView(viewModel));
        }
        private ImportViewModel ReturnDataView(ImportViewModel vm)
        {

            try
            {
                vm.ListItem = CheckListItems(vm.SearchDate);
                vm.ListData = new DataViewModel() { DataTableHeader = new Models.Entity.DataModel(), DataTableBody = _importData.FindDataOfLastTimeInit(vm.SearchDate) };
            }
            catch(Exception ex)
            {
                var error = ex.Message;
                //var GetType = ex.GetType;
                TempData["error"] = "途中にエラーが発生しました!";
                return new ImportViewModel();
            }
            
            return vm;
        }

        [HttpGet]
        public IActionResult SearchData()
        {
            var viewModel = new ImportViewModel(); 
            try
            {
                viewModel.ListItem = CheckListItems(viewModel.SearchDate);
                viewModel.ListData = new DataViewModel() { DataTableHeader = new Models.Entity.DataModel(), DataTableBody = _importData.FindDataOfLastTime(viewModel) };

            }
            catch (Exception ex)
            {
                return RedirectToAction("ImportData", new ImportViewModel());
            }

            return View("ImportData", viewModel);
        }

        [HttpPost]
        public IActionResult SearchData(ImportViewModel viewModel)
        {
            try
            {
                if (!DateTime.TryParse(viewModel.ItemValue, out DateTime dateValue))
                {
                    return View("ImportData", new ImportViewModel());
                }
                if(viewModel.SearchDate <= DateMin || viewModel.SearchDate >= DateMax)
                {
                    return View("ImportData", new ImportViewModel());
                }
                viewModel.ListItem = CheckListItems(viewModel.SearchDate);
                viewModel.ListData = new DataViewModel() { DataTableHeader = new Models.Entity.DataModel(), DataTableBody = _importData.FindDataOfLastTime(viewModel)};
                
            }
            catch (Exception ex)
            {
                return RedirectToAction("ImportData", new ImportViewModel());
            }

            return View("ImportData", viewModel);
        }

        [HttpPost]
        public IActionResult DeleteData()
        {
            var viewModel = new ImportViewModel();
            var date = DateTime.Today;
            var rowsAffected = _importData.DeleteData(date);
            if (rowsAffected > 0)
            {
                ComShowSuccessMessage("本日のデータは削除が完了しました！");
                return View("ImportData", ReturnDataView(viewModel));
            }
            else if(rowsAffected == 0)
            {
               ComShowErrorMessage("削除に失敗しました！");
                return View("ImportData", ReturnDataView(viewModel));
            }
            else
            {
                ComShowErrorMessage("本日はデータが存在していませんので削除に失敗しました！");
                return View("ImportData", ReturnDataView(viewModel));
            }
        }
        private List<object> ReaderData(IFormFile file, ref List<object> result, ref List<Bu_MastarModel> bu_Mastars)
        {

            var fileName = file.FileName;
            try
            {
                using (var reader = new StreamReader(file.OpenReadStream()))
                {
                    var lineNo = 1;
                    int lengh = 0;
                    int Position = 1;
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        if (lineNo <= 1)
                        {
                            lineNo++;
                            continue;
                        }
                        if (line == "\u001a")
                        {
                            break;
                        }
                        // 1行ずつ読み込む
                        lengh = line.Length;
                        // 各行つずつチェック
                        if ((lengh + 1) < 240)
                        {
                            result = new List<object>();
                            var errmess = "データのフォーマットが正しくないため、読み込めませんでした";
                            var status = "NG";
                            result.Add(status);
                            result.Add(fileName);
                            result.Add(errmess);
                            result.Add(lineNo.ToString());
                            break;
                        }
                        //string iDate = "2005-05-05";
                        //DateTime oDate = DateTime.Parse(iDate);
                        // 1件目(左分)
                        dataModel = new DataModel()
                        {
                            WAYMD = new DateTime(Convert.ToInt32("20" + line.Substring(17, 2)), Convert.ToInt32(line.Substring(19, 2)), Convert.ToInt32(line.Substring(21, 2)), 00, 00, 00),
                            SEQ = Util.NullToBlank((object)line.Substring(23, 4)),
                            KATASIKI = Util.NullToBlank(line.Substring(27, 12)),
                            MEISHO = Util.NullToBlank(line.Substring(39, 1)),
                            FILLER1 = Util.NullToBlank(line.Substring(40, 2)),
                            OPT = Util.NullToBlank(line.Substring(42, 3)),
                            JIKU = Util.NullToBlank(line.Substring(45, 8)),
                            FILLER2 = Util.NullToBlank(line.Substring(53, 2)),
                            DAI = Util.NullToBlank((object)line.Substring(55, 2)),
                            MC = Util.NullToBlank(line.Substring(57, 2)),
                            SIMUKE = Util.NullToBlank(line.Substring(59, 1)),
                            E0 = Util.NullToBlank(line.Substring(60, 2)),
                            BUBAN = Util.NullToBlank(line.Substring(62, 17)),
                            TANTO = Util.NullToBlank(line.Substring(79, 1)),
                            GR = Util.NullToBlank(line.Substring(80, 2)),
                            KIGO = Util.NullToBlank(line.Substring(82, 4)),
                            MAKR = Util.NullToBlank(line.Substring(86, 4)),
                            KOSUU = Util.NullToBlank((object)line.Substring(90, 3)),
                            KISYU = Util.NullToBlank(line.Substring(93, 2)),
                            MEWISYO = Util.NullToBlank(line.Substring(95, 18)),
                            FYMD = new DateTime(Convert.ToInt32("20" + line.Substring(113, 2)), Convert.ToInt32(line.Substring(115, 2)), Convert.ToInt32(line.Substring(117, 2)), 00, 00, 00), // line.Substring(113, 6),
                            SEIHINCD = Util.NullToBlank(line.Substring(119, 3)),
                            SEHINJNO = Util.NullToBlank(line.Substring(122, 6)),
                            FileName = Util.NullToBlank(file.FileName),
                            LineNumber = lineNo,
                            Position = Position++
                        };
                        result.Add(dataModel);
                        // Dictionaryの更新
                        var BUBAN = Util.NullToBlank(line.Substring(62, 17));
                        var MEWISYO = Util.NullToBlank(line.Substring(95, 18));
                        var KIGO = Util.NullToBlank(line.Substring(82, 4));
                        if (BUBAN != string.Empty)
                        {
                            CheckBuMastar(BUBAN, MEWISYO, KIGO, ref bu_Mastars);
                        }

                        // 2件目(右分)
                        dataModel = new DataModel()
                        {
                            WAYMD = new DateTime(Convert.ToInt32("20" + line.Substring(128, 2)), Convert.ToInt32(line.Substring(130, 2)), Convert.ToInt32(line.Substring(132, 2)), 00, 00, 00),
                            SEQ = Util.NullToBlank((object)line.Substring(134, 4)),
                            KATASIKI = Util.NullToBlank(line.Substring(138, 12)),
                            MEISHO = Util.NullToBlank(line.Substring(150, 1)),
                            FILLER1 = Util.NullToBlank(line.Substring(151, 2)),
                            OPT = Util.NullToBlank(line.Substring(153, 3)),
                            JIKU = Util.NullToBlank(line.Substring(156, 8)),
                            FILLER2 = Util.NullToBlank(line.Substring(164, 2)),
                            DAI = Util.NullToBlank((object)line.Substring(166, 2)),
                            MC = Util.NullToBlank(line.Substring(168, 2)),
                            SIMUKE = Util.NullToBlank(line.Substring(170, 1)),
                            E0 = Util.NullToBlank(line.Substring(171, 2)),
                            BUBAN = Util.NullToBlank(line.Substring(173, 17)),
                            TANTO = Util.NullToBlank(line.Substring(190, 1)),
                            GR = Util.NullToBlank(line.Substring(191, 2)),
                            KIGO = Util.NullToBlank(line.Substring(193, 4)),
                            MAKR = Util.NullToBlank(line.Substring(197, 4)),
                            KOSUU = Util.NullToBlank((object)line.Substring(201, 3)),
                            KISYU = Util.NullToBlank(line.Substring(204, 2)),
                            MEWISYO = Util.NullToBlank(line.Substring(206, 18)),
                            FYMD = new DateTime(Convert.ToInt32("20" + line.Substring(224, 2)), Convert.ToInt32(line.Substring(226, 2)), Convert.ToInt32(line.Substring(228, 2)), 00, 00, 00), // line.Substring(224, 6),
                            SEIHINCD = Util.NullToBlank(line.Substring(230, 3)),
                            SEHINJNO = Util.NullToBlank(line.Substring(233, 6)),
                            FileName = Util.NullToBlank(file.FileName),
                            LineNumber = lineNo,
                            Position = Position++
                        };
                        lineNo++;
                        result.Add(dataModel);
                        // Dictionaryの更新
                        BUBAN = Util.NullToBlank(line.Substring(173, 17));
                        MEWISYO = Util.NullToBlank(line.Substring(206, 18));
                        KIGO = Util.NullToBlank(line.Substring(193, 4));
                        if (BUBAN != string.Empty)
                        {
                            CheckBuMastar(BUBAN, MEWISYO, KIGO, ref bu_Mastars);
                        }
                    }

                }
            }
            catch
            {
                throw;
            }

            return result;
        }
        private void CheckBuMastar(string buBan, string meWiSyo, string kigo, ref List<Bu_MastarModel> bu_Mastars)
        {
            if (bu_Mastars == null)
            {
                throw new Exception("部品番号が存在しません");
            }
            //var result = bu_Mastars.First(s => s.BUBAN == buBan);
            var item1 = (from item in bu_Mastars
                         where item.BUBAN == buBan
                         select item).FirstOrDefault();

            StringComparison comp = StringComparison.OrdinalIgnoreCase;
            bool check1 = meWiSyo.Contains(ExportExcelController.FL00R_ASSY, comp);
            bool check2 = meWiSyo.Contains(ExportExcelController.FRAME_ASSY, comp);
            if (item1 == null)
            {
                if (check1)
                {
                    bu_Mastars.Add(new Bu_MastarModel() { BUBAN = buBan, KIGO = kigo, MEWISYO = meWiSyo, Nyusu = ExportExcelController.PALETNO_FLOOR_ASSY, KatakanaName = KATAKANA_NAME_FLOOR_ASSY });
                }
                else if (check2)
                {
                    bu_Mastars.Add(new Bu_MastarModel() { BUBAN = buBan, KIGO = kigo, MEWISYO = meWiSyo, Nyusu = ExportExcelController.PALETNO_FLAME_ASSY, KatakanaName = KATAKANA_NAME_FLAME_ASSY });
                }
                else
                {
                    bu_Mastars.Add(new Bu_MastarModel() { BUBAN = buBan, KIGO = kigo, MEWISYO = meWiSyo, Nyusu = 9999, KatakanaName = KATAKANA_NAME_HOKANO });
                }
            }
            return;
        }

        [HttpPost]
        public JsonResult GetDropListAjax(DateTime date)
        {
            var items = new Tuple<List<string>, List<string>>(new List<string>(), new List<string>());
            List<SelectListItem> droplists = new List<SelectListItem>();
            try
            {
                items = _importData.FindDropList(date);
                var i = 0;
                bool isSelected = true;
                foreach (var item in items.Item1)
                {
                    if (i == 0) { isSelected = true; }
                    else { isSelected = false; }
                    droplists.Add(new SelectListItem()
                    {
                        Text = item,
                        Value = i.ToString(),
                        Selected = isSelected
                    });
                    i++;
                }
                for (int j = 0; j < droplists.Count; j++)
                {
                    droplists[j].Value = items.Item2[j];
                }
                // 更新処理
                ListItems = items.Item2;
            }
            catch (Exception ex)
            {
                ErrorInfor.DebugWriteLineError(ex);
                Debug.WriteLine("========================================================================================");
                return Json(new { StatusCode = false, Mess = "error" });
            }

            return Json(new { StatusCode = true, droplists });
        }

        private List<SelectListItem> CheckListItems(DateTime date)
        {
            var items = new Tuple<List<string>, List<string>>(new List<string>(), new List<string>());
            List<SelectListItem> droplists = new List<SelectListItem>();
            try
            {
                items = _importData.FindDropList(date);
                var i = 0;
                bool isSelected = true;
                foreach (var item in items.Item1)
                {
                    if (i == 0) { isSelected = true; }
                    else { isSelected = false; }
                    droplists.Add(new SelectListItem()
                    {
                        Text = item,
                        Value = i.ToString(),
                        Selected = isSelected
                    });
                    i++;
                }
                for (int j = 0; j < droplists.Count; j++)
                {
                    droplists[j].Value = items.Item2[j];
                }
            }
            catch (Exception ex)
            {
                ErrorInfor.DebugWriteLineError(ex);
                Debug.WriteLine("========================================================================================");
                throw;
            }

            return droplists;
        }
    }
}
