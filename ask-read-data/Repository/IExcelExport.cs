using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using ask_read_data.Models;

namespace ask_read_data.Repository
{
    public interface IExcelExport
    {
        DataTable GetFloor_Flame_Assy(DateTime dateTime, string BubanType);
    }
}
