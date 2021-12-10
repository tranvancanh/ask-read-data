using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using ask_read_data.Areas.Admin.Commons;
using ask_read_data.Models;
using ask_read_data.Repository;
using ask_tzn_funamiKD.Commons;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace ask_read_data.Controllers
{
    [Authorize]
    public class ImportDataController : Controller
    {
        private readonly IImportData _importData;
        private DataModel dataModel;
        public ImportDataController(IImportData importData, DataModel dataModel)
        {
            this._importData = importData;
            this.dataModel = dataModel;
        }

        [HttpGet]
        public IActionResult ImportData()
        {
            return View();
        }

        [HttpPost]
        public IActionResult ImportData(List<IFormFile> FileUpload)
        {
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
            var result = new List<object>();
            var dicts = new Dictionary<string, string>();
            var files = FileUpload;
            var fileNamee = "";
            if (files == null || files.Count <= 0)
            {
                ViewData["erroremess"] = "ファイル内にデータが存在しません";
                return View();
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
                            return View();
                        }

                        //1ファイルデータずつ読み取り
                        ReaderData(file, ref result, ref dicts);
                        var firstPosition = result.FirstOrDefault().ToString();
                        if (firstPosition == "NG")
                        {
                            var fileName = result[1];
                            var lineNo = result[3];
                            ViewData["successmess"] = null;
                            ViewData["erroremess"] = fileName.ToString() + "： ファイル読み込み中にエラーが発生しました" + " | lineNo: " + lineNo.ToString();
                            return View();
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                ViewData["successmess"] = null;
                ViewData["erroremess"] = ex.Message + $@" | file name: {fileNamee}";
                return View();
            }
            // go to service
            ResponResult res = _importData.ImportDataDB(result, Claims);
            if (res.Status != "OK")
            {
                ViewData["successmess"] = null;
                ViewData["erroremess"] = res.Resmess;
                return View();
            }

            // ViewData["successmess"] = "ファイル内にデータが保存されました";
            ViewData["erroremess"] = null;
            ViewData["successmess"] = res.Resmess;
            
            return View();
        }

        private List<object> ReaderData(IFormFile file, ref List<object> result, ref Dictionary<string, string> dicts)
        {

            var fileName = file.FileName;
            try
            {
                using (var reader = new StreamReader(file.OpenReadStream()))
                {
                    var lineNo = 1;
                    int lengh = 0;
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
                        // 1件目(左分)
                        dataModel = new DataModel()
                        {
                            WAYMD = new DateTime(Convert.ToInt32("20" + line.Substring(17, 2)), Convert.ToInt32(line.Substring(19, 2)), Convert.ToInt32(line.Substring(21, 2)), 00, 00, 00),
                            SEQ = Util.NullToBlank((object)line.Substring(23, 4)),
                            KATASIKI = line.Substring(27, 12),
                            MEISHO = line.Substring(39, 1),
                            FILLER1 = line.Substring(40, 2),
                            OPT = line.Substring(42, 3),
                            JIKU = line.Substring(45, 8),
                            FILLER2 = line.Substring(53, 2),
                            DAI = Util.NullToBlank((object)line.Substring(55, 2)),
                            MC = line.Substring(57, 2),
                            SIMUKE = line.Substring(59, 1),
                            E0 = line.Substring(60, 2),
                            BUBAN = line.Substring(62, 17),
                            TANTO = line.Substring(79, 1),
                            GR = line.Substring(80, 2),
                            KIGO = line.Substring(82, 4),
                            MAKR = line.Substring(86, 4),
                            KOSUU = Util.NullToBlank((object)line.Substring(90, 3)),
                            KISYU = line.Substring(93, 2),
                            MEWISYO = line.Substring(95, 18),
                            FYMD = new DateTime(Convert.ToInt32("20" + line.Substring(113, 2)), Convert.ToInt32(line.Substring(115, 2)), Convert.ToInt32(line.Substring(117, 2)), 00, 00, 00), // line.Substring(113, 6),
                            SEIHINCD = line.Substring(119, 3),
                            SEHINJNO = line.Substring(122, 6),
                            FileName = file.FileName,
                            LineNumber = lineNo
                        };
                        result.Add(dataModel);
                        // Dictionaryの更新


                        // 2件目(右分)
                        dataModel = new DataModel()
                        {
                            WAYMD = new DateTime(Convert.ToInt32("20" + line.Substring(128, 2)), Convert.ToInt32(line.Substring(130, 2)), Convert.ToInt32(line.Substring(132, 2)), 00, 00, 00),
                            SEQ = Util.NullToBlank((object)line.Substring(134, 4)),
                            KATASIKI = line.Substring(138, 12),
                            MEISHO = line.Substring(150, 1),
                            FILLER1 = line.Substring(151, 2),
                            OPT = line.Substring(153, 3),
                            JIKU = line.Substring(156, 8),
                            FILLER2 = line.Substring(164, 2),
                            DAI = Util.NullToBlank((object)line.Substring(166, 2)),
                            MC = line.Substring(168, 2),
                            SIMUKE = line.Substring(170, 1),
                            E0 = line.Substring(171, 2),
                            BUBAN = line.Substring(173, 17),
                            TANTO = line.Substring(190, 1),
                            GR = line.Substring(191, 2),
                            KIGO = line.Substring(193, 4),
                            MAKR = line.Substring(197, 4),
                            KOSUU = Util.NullToBlank((object)line.Substring(201, 3)),
                            KISYU = line.Substring(204, 2),
                            MEWISYO = line.Substring(206, 18),
                            FYMD = new DateTime(Convert.ToInt32("20" + line.Substring(224, 2)), Convert.ToInt32(line.Substring(226, 2)), Convert.ToInt32(line.Substring(228, 2)), 00, 00, 00), // line.Substring(224, 6),
                            SEIHINCD = line.Substring(230, 3),
                            SEHINJNO = line.Substring(233, 6),
                            FileName = file.FileName,
                            LineNumber = lineNo
                        };
                        lineNo++;
                        result.Add(dataModel);
                        // Dictionaryの更新


                    }

                }
            }
            catch
            {
                throw;
            }

            return result;
        }
    }
}
