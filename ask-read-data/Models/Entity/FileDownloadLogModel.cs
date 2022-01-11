using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace ask_read_data.Models.Entity
{
    public class FileDownloadLogModel
    {
        [StringLength(50, ErrorMessage = "50文字以内で入力してください", MinimumLength = 1)]
        [DataType(DataType.Text)]
        [Display(Name = "部品番号")]
        public string BubanMeiType { get; set; }

        [Display(Name = "最終ダウンロード日時")]
        [DataType(DataType.Date)]
        public DateTime LastDownloadDateTime { get; set; }

        public int Position { get; set; }

        [Display(Name = "パレット連番")]
        public int ParetoRenban { get; set; }

        [Display(Name = "ダウンロード詳細時間")]
        public DateTime DownloadDateTime { get; set; }

        [StringLength(50, ErrorMessage = "50文字以内で入力してください", MinimumLength = 0)]
        [DataType(DataType.Text)]
        public string udownload { get; set; }

        public FileDownloadLogModel()
        {
            BubanMeiType = Controllers.ExportExcelController.FL00R_ASSY;
            LastDownloadDateTime = new DateTime(1900, 01, 01, 00, 00, 00);
            Position = 0;
            ParetoRenban = 0;
            DownloadDateTime = new DateTime(1900, 01, 01, 00, 00, 00);
            udownload = "usertozan";
        }
    }
}
