using System;
using System.Collections.Generic;
using System.Security.Claims;
using ask_read_data.Models;
using ask_read_data.Models.Entity;
using ask_read_data.Models.ViewModel;

namespace ask_read_data.Repository
{
    public interface IImportData
    {
        ResponResult ImportDataDB(List<object> datas, List<Claim> claims, List<Bu_MastarModel> buMastars);
        Tuple<List<string>, List<string>> FindDropList(DateTime date);
        List<DataModel> FindDataOfLastTimeInit(DateTime date);
        List<DataModel> FindDataOfLastTime(ImportViewModel viewModel);
        int DeleteData(DateTime date);
    }
}
