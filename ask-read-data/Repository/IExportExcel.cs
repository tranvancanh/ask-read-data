using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using ask_read_data.Models;
using ask_read_data.Models.Entity;
using ask_read_data.Models.ViewModel;

namespace ask_read_data.Repository
{
    public interface IExportExcel
    {
        (DataTable, DataTable) GetFloor_Flame_Assy(ExportExcelViewModel modelRequset, string bubanType);
        int RecordDownloadHistory(ref DataTable dataTable, string bubanType, List<Claim> Claims);
        Tuple<int, int> FindPositionParetoRenban(DateTime date, string bubanType);
        List<DataModel> FindRemainingDataOfLastTime(ExportExcelViewModel viewModel);
        List<FileDownloadLogModel> FindDownloadHistory(DateTime date);
    }
}
