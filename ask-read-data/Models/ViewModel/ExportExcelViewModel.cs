using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace ask_read_data.Models.ViewModel
{
    public class ExportExcelViewModel
    {
        [DataType(DataType.Date)]
        public DateTime Floor_Assy { get; set; }

        [DataType(DataType.Date)]
        public DateTime Flame_Assy { get; set; }

        public ExportExcelViewModel()
        {
            Floor_Assy = Convert.ToDateTime(DateTime.Today.ToString("yyyy-MM-dd 00:00:00"));
            Flame_Assy = Convert.ToDateTime(DateTime.Today.ToString("yyyy-MM-dd 00:00:00"));
        }
    }
}
