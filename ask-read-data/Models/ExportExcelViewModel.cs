using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ask_read_data.Models
{
    public class ExportExcelViewModel
    {
        // --------------------Page Header --------------------  ///////
        public DateTime Yotebi { get; set; }

        // -------------------- Table --------------------     ////////
        public DataModel DataTableHeader { get; set; } 
        public List<DataModel> DataTableBody { get; set; }

        public ExportExcelViewModel()
        {
            Yotebi = DateTime.Today;
            DataTableHeader  = new DataModel();
            DataTableBody = new List<DataModel>();
        }

    }
}
