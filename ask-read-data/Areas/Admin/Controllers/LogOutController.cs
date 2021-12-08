using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace ask_read_data.Areas.Admin.Controllers
{
    [Authorize]
    public class LogOutController : Controller
    {
        [HttpGet]
        public async Task<IActionResult> LogOut()
        {
            // 認証クッキーをレスポンスから削除
            await HttpContext.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme);

            // ログイン画面にリダイレクト
            return RedirectToAction("Login", "Login");
        }
    }
}
