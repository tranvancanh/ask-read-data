using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ask_read_data.Models
{
    public class DataViewModel1
    {
        // --------------------Page Header --------------------  ///////
        public DateTime ImportDate { get; set; }

        // -------------------- Table --------------------     ////////
        public DataModel DataTableHeader { get; set; } 
        public List<DataModel> DataTableBody { get; set; }

        public DataViewModel1()
        {
            ImportDate = DateTime.Today;
            DataTableHeader  = new DataModel();
            DataTableBody = new List<DataModel>();
        }

    }
}
