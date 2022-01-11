using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc.Rendering;

namespace ask_read_data.Models.ViewModel
{
    public class ExportExcelViewModel
    {
        [DataType(DataType.Date)]
        public DateTime Floor_Assy { get; set; }

        public int Floor_Position { get; set; }
        public int Floor_ParetoRenban { get; set; }

        [DataType(DataType.Date)]
        public DateTime Flame_Assy { get; set; }

        public int Flame_Position { get; set; }
        public int Flame_ParetoRenban { get; set; }

        /////////////////////////////////////////////   Search分   //////////////////////////////////////

        [DataType(DataType.Date)]
        public DateTime SearchDate { get; set; }

        public DownloadHistoryViewModel ListData { get; set; }

        public ExportExcelViewModel()
        {
            Floor_Assy = Convert.ToDateTime(DateTime.Today.ToString("yyyy-MM-dd 00:00:00"));
            Floor_Position = 0;
            Floor_ParetoRenban = 0;
            Flame_Assy = Convert.ToDateTime(DateTime.Today.ToString("yyyy-MM-dd 00:00:00"));
            Flame_Position = 0;
            Flame_ParetoRenban = 0;
            SearchDate = Convert.ToDateTime(DateTime.Today.ToString("yyyy-MM-dd 00:00:00"));
        }
    }
}
