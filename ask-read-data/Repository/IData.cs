using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ask_read_data.Models;
using ask_read_data.Models.Entity;

namespace ask_read_data.Repository
{
    public interface IData
    {
        List<DataModel> GetAll2000DataImport();
        List<DataModel> SearchDataImport(DateTime dateTime);
    }
}
