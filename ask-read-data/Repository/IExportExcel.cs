using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ask_read_data.Models;

namespace ask_read_data.Repository
{
    public interface IExportExcel
    {
        List<DataModel> GetAll2000DataImport();
        List<DataModel> SearchDataImport(DateTime dateTime);
    }
}
