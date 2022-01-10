using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ask_read_data.Models
{
    public class TemporaryData
    {
        public string Position { get; set; }
        public DateTime CreateDateTime { get; set; }

        public TemporaryData()
        {
            Position = "Position";
            CreateDateTime = new DateTime(1900, 01, 01 , 00, 00, 00);
        }
    }
}
