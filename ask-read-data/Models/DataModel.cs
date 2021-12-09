using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ask_read_data.Models
{
    public class DataModel
    {
        public DateTime WAYMD { get; set; }
        public string SEQ { get; set; }
        public string KATASIKI { get; set; }
        public string MEISHO { get; set; }
        public string FILLER1 { get; set; }
        public string OPT { get; set; }
        public string JIKU { get; set; }
        public string FILLER2 { get; set; }
        public string DAI { get; set; }
        public string MC { get; set; }
        public string SIMUKE { get; set; }
        public string E0 { get; set; }
        public string BUBAN { get; set; }
        public string TANTO { get; set; }
        public string GR { get; set; }
        public string KIGO { get; set; }
        public string MAKR { get; set; }
        public string KOSUU { get; set; }
        public string KISYU { get; set; }
        public string MEWISYO { get; set; }
        public string FYMD { get; set; }
        public string SEIHINCD { get; set; }
        public string SEHINJNO { get; set; }

        public DataModel()
        {
            WAYMD = new DateTime(1900, 01, 01, 00, 00, 00);
            SEQ = "";
            KATASIKI = "";
            MEISHO = "";
            FILLER1 = "";
            OPT = "";
            JIKU = "";
            FILLER2 = "";
            DAI = "";
            MC = "";
            SIMUKE = "";
            E0 = "";
            BUBAN = "";
            TANTO = "";
            GR = "";
            KIGO = "";
            MAKR = "";
            KOSUU = "";
            KISYU = "";
            MEWISYO = "";
            FYMD = "";
            SEIHINCD = "";
            SEHINJNO = "";

        }
    }
}