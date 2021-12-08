using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using ask_read_data.Areas.Admin.Models;
using ask_read_data.Areas.Admin.Repository;

namespace ask_read_data.Areas.Admin.Controllers
{
    //[Authorize(Roles ="1")]
    [Area("Admin")]
    [Authorize]
    public class UserInfoController : Controller
    {
        public readonly IUser _user;
        public UserInfoController(IUser user)
        {
            this._user = user;
        }

        [HttpGet]
        public IActionResult Index()
        {
            string UserName = User.Claims.Where(c => c.Type == ClaimTypes.Name).First().Value;
            string Password = User.Claims.Where(c => c.Type == "Password").First().Value;
            UserViewModel user = _user.GetUser(UserName, Password);
            return View(user);
        }
    }
}
