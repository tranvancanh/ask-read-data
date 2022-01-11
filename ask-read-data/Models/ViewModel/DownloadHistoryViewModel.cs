using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ask_read_data.Models.Entity;

namespace ask_read_data.Models.ViewModel
{
    public class DownloadHistoryViewModel
    {
        // -------------------- Table --------------------     ////////
        public FileDownloadLogModel DataTableHeader { get; set; }
        public List<FileDownloadLogModel> DataTableBody { get; set; }

        public DownloadHistoryViewModel()
        {
            DataTableHeader = new FileDownloadLogModel();
            DataTableBody = new List<FileDownloadLogModel>();
        }
    }
}
