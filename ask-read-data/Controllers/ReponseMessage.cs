using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace ask_read_data.Commons
{
    public class ReponseMessage : Controller
    {
        public void ComShowErrorMessage(string msg)
        {
            ViewData["ComShowErrorMessage"] = msg;
            return;
        }
        public void ComShowSuccessMessage(string msg)
        {
            ViewData["ComShowSuccessMessage"] = msg;
            return;
        }
    }
}
