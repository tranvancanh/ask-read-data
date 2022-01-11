using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace ask_read_data.Commons
{
    public class ReponseMessage : Controller
    {
        [TempData]
        public string ComShowSuccessMessage1 { get; set; }

        [TempData]
        public string ComShowErrorMessage1 { get; set; }

        public void ComShowSuccessMessage(string msg)
        {
            TempData["ComShowSuccessMessage"] = msg;
            return;
        }
        public void ComShowErrorMessage(string msg)
        {
            TempData["ComShowErrorMessage"] = msg;
            return;
        }

    }
}
