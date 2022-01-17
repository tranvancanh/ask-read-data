using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc.Rendering;

namespace ask_read_data.Models.ViewModel
{
    public class ImportViewModel
    {
        public string ItemValue { get; set; }

        public List<SelectListItem> ListItem { get; set; }

        public string BubanType { get; set; }

        [DataType(DataType.Date)]
        public DateTime SearchDate { get; set; }

        public DataViewModel ListData { get; set; }

        public ImportViewModel()
        {
            ItemValue = DateTime.Today.ToString("yyyy") + "/" + DateTime.Today.ToString("MM") + "/" + DateTime.Today.ToString("dd");
            BubanType = "ALL";
            SearchDate = Convert.ToDateTime(DateTime.Today.ToString("yyyy-MM-dd 00:00:00"));
            ListData = new DataViewModel();
        }
    }
}
