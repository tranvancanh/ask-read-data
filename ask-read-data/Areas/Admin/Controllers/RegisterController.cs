using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace ask_read_data.Areas.Admin.Controllers
{
    [Area("Admin")]
    public class RegisterController : Controller
    {
        public IActionResult Register()
        {
            return View();
        }
    }
}
