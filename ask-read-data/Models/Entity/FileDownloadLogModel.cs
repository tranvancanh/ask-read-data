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
        public string BubanMeiType { get; set; }
        public DateTime LastImportDateTime { get; set; }
        public int Position { get; set; }
        public DateTime DownloadDateTime { get; set; }

        [StringLength(50, ErrorMessage = "50文字以内で入力してください", MinimumLength = 0)]
        [DataType(DataType.Text)]
        public string udownload { get; set; }

        public FileDownloadLogModel()
        {
            BubanMeiType = Controllers.ExportExcelController.FL00R_ASSY;
            LastImportDateTime = new DateTime(1900, 01, 01, 00, 00, 00);
            Position = 0;
            DownloadDateTime = new DateTime(1900, 01, 01, 00, 00, 00);
            udownload = "usertozan";
        }
    }
}
