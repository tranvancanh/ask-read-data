using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ask_read_data.Models.ViewModel
{
    public class DataViewModel
    {
        // -------------------- Table --------------------     ////////
        public DataModel DataTableHeader { get; set; }
        public List<DataModel> DataTableBody { get; set; }

        public DataViewModel()
        {
            DataTableHeader = new DataModel();
            DataTableBody = new List<DataModel>();
        }
    }
}
