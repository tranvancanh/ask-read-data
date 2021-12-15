using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using ask_read_data.Models;

namespace ask_read_data.Repository
{
    public interface IExportExcel
    {
        (DataTable, DataTable) GetFloor_Flame_Assy(DateTime dateTime, string bubanType);
        int RecordDownloadHistory(ref DataTable dataTable, string bubanType, List<Claim> Claims);
    }
}
