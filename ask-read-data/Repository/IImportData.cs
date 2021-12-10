using System.Collections.Generic;
using System.Security.Claims;
using ask_read_data.Models;

namespace ask_read_data.Repository
{
    public interface IImportData
    {
        ResponResult ImportDataDB(List<object> datas, List<Claim> claims);
    }
}
