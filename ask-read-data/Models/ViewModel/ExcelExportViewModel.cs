using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ask_read_data.Models.ViewModel
{
    public class ExcelExportViewModel
    {
        public DateTime Floor_Assy { get; set; }
        public DateTime Flame_Assy { get; set; }

        public ExcelExportViewModel()
        {
            Floor_Assy = Convert.ToDateTime(DateTime.Today.ToString("yyyy-MM-dd 00:00:00"));
            Flame_Assy = Convert.ToDateTime(DateTime.Today.ToString("yyyy-MM-dd 00:00:00"));
        }
    }
}
