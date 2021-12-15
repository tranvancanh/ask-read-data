using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ask_read_data.Models.Entity
{
    public class FileImportLogModel
    {
        public DateTime CreateDateTime { get; set; }
        public string FileName { get; set; }
        public int TotalLine { get; set; }
        public int MaxPosition { get; set; }
        public string Createby { get; set; }

        public FileImportLogModel()
        {
            CreateDateTime = Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            FileName = "Tozanadmin";
            TotalLine = 0;
            MaxPosition = 0;
            Createby = "Tozanadmin";
        }
    }
}