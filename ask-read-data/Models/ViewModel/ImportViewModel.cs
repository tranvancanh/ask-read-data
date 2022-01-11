using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace ask_read_data.Models.ViewModel
{
    public class ImportViewModel
    {
        public string ItemValue { get; set; }

        public string BubanType { get; set; }

        [DataType(DataType.Date)]
        public DateTime SearchDate { get; set; }

        public DataViewModel ListData { get; set; }

        public ImportViewModel()
        {
            BubanType = "ALL";
            SearchDate = Convert.ToDateTime(DateTime.Today.ToString("yyyy-MM-dd 00:00:00"));
            ListData = new DataViewModel();
        }
    }
}
