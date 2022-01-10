using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using ask_read_data.Areas.Admin.Commons;

namespace ask_read_data.Models
{
    public class DataModel
    {

        [Display(Name = "溶組予定日")]
        [DataType(DataType.Date)]
        public DateTime WAYMD { get; set; }

        public int SEQ { get; set; }

        [Display(Name = "型式番号")]
        public string KATASIKI { get; set; }
        public string MEISHO { get; set; }
        public string FILLER1 { get; set; }
        public string OPT { get; set; }

        [Display(Name = "時区")]
        public string JIKU { get; set; }

        public string FILLER2 { get; set; }

        public int DAI { get; set; }

        public string MC { get; set; }

        public string SIMUKE { get; set; }

        public string E0 { get; set; }

        [Display(Name = "部品番号")]
        public string BUBAN { get; set; }

        public string TANTO { get; set; }

        public string GR { get; set; }

        [Display(Name = "部品略式記号")]
        public string KIGO { get; set; }

        [Display(Name = "取引先")]
        public string MAKR { get; set; }

        public int KOSUU { get; set; }

        public string KISYU { get; set; }

        [Display(Name = "部品名称")]
        public string MEWISYO { get; set; }

        [Display(Name = "発送予定日")]
        [DataType(DataType.Date)]
        public DateTime FYMD { get; set; }

        public string SEIHINCD { get; set; }

        public string SEHINJNO { get; set; }

        [Display(Name = "ファイル名")]
        public string FileName { get; set; }

        public int LineNumber { get; set; }

        [Display(Name = "Position")]
        public int Position { get; set; }

        public string UpBy { get; set; }

        [Display(Name = "作成時間")]
        [DataType(DataType.DateTime)]
        public DateTime CreateDateTime { get; set; }

        public string CreateBy { get; set; }

        public DataModel()
        {
            WAYMD = new DateTime(1900, 01, 01, 00, 00, 00);
            SEQ = 0;
            KATASIKI = "";
            MEISHO = "";
            FILLER1 = "";
            OPT = "";
            JIKU = "";
            FILLER2 = "";
            DAI = 0;
            MC = "";
            SIMUKE = "";
            E0 = "";
            BUBAN = "";
            TANTO = "";
            GR = "";
            KIGO = "";
            MAKR = "";
            KOSUU = 0;
            KISYU = "";
            MEWISYO = "";
            FYMD = new DateTime(1900, 01, 01, 00, 00, 00);
            SEIHINCD = "";
            SEHINJNO = "";
            FileName = "tozan.abc";
            LineNumber = 0;
            UpBy = "";     //(new UserInfor()).UserInfo().UserName;
            CreateDateTime = new DateTime(1900, 01, 01, 00, 00, 00);
            CreateBy = ""; // new UserInfor().UserInfo().UserName;

        }
    }
}