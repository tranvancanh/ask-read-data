using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using ask_read_data.Areas.Admin.Models;
using ask_read_data.Areas.Admin.Repository;
using ask_tzn_funamiKD.Commons;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using System.Security.Claims;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authentication;
using Microsoft.IdentityModel.Logging;
using Microsoft.IdentityModel.Tokens;
using System.Text;
using System.IdentityModel.Tokens.Jwt;
using Microsoft.Extensions.Configuration;

namespace ask_read_data.Areas.Admin.Controllers
{
    [Area(areaName:"Admin")]
    //[Route("Admin/Account/Login")]
    [AllowAnonymous]
    public class LoginController : Controller
    {
        private readonly ILoginViewModel _login;
        private readonly IRegisterService _registerService;
        //////////////////////     Inject Session  /////////////////////
        private readonly ISession Session;
        private readonly IConfiguration _configuration;
        private UserViewModel userViewModel;
        public readonly IUser _user;
        public LoginController(ILoginViewModel login, IRegisterService registerService, IConfiguration configuration, UserViewModel userViewModel, IUser user)
        {
            // IHttpContextAccessor httpContextAccessor
            this._login = login;
            this._registerService = registerService;
            //this.Session = httpContextAccessor.HttpContext.Session;
            this._configuration = configuration;
            this.userViewModel = userViewModel;
            this._user = user;
        }
        const string UserName = "";
        /// <summary>
        /// //////////////////////     ============================================================================    /////////////////////
        /// </summary>
        private void CheckUserLogin(LoginViewModel loginModel)
        {
            string UserName = Util.NullToBlank(loginModel.UserName);
            string Password = Util.NullToBlank(loginModel.Password);

            ///////////////// Check UserName ////////////////////////////////////
            if (string.IsNullOrWhiteSpace(UserName) == true)
                throw new Exception("ユーザーIDが入力してください");
            if (UserName.Length < 2 || UserName.Length > 50)
                throw new Exception("ユーザーIDは2~50文字以内で入力してください");
            ///////////////// Check Password ////////////////////////////////////
            if (string.IsNullOrWhiteSpace(Password) == true)
                throw new Exception("パスワードが入力してください");
            if (Password.Length < 2 || Password.Length > 50)
                throw new Exception("パスワードは2~50文字以内で入力してください");
        }
        [HttpGet]
        public IActionResult Login(string returnUrl)
        {
            ViewBag.ReturnUrl = returnUrl;
            return View("Login");
        }

        [HttpPost]
        [AllowAnonymous]
        [ValidateAntiForgeryToken]
        [ActionName("Login")]
        public async Task<IActionResult> Login([FromForm]LoginViewModel loginModel, string returnUrl)
        {
            //ContactDBEntities db = new ContactDBEntities();
            try
            {
                CheckUserLogin(loginModel);
            }
            catch (Exception ex)
            {
                ViewData.Add("errormess", ex.Message);
                return View("Login");
            }
            bool isLogin = _login.UserLogin(loginModel);
            if (isLogin == true)
            {
                
                userViewModel = _user.GetUser(loginModel.UserName, loginModel.Password);

                // HttpContext.Session.SetString("UserName", loginModel.UserName);
                ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                if(userViewModel == null)
                {
                    ViewData.Add("errormess", "ユーザーIDまたはパスワードに誤りがあります");
                    return View("Login");
                }
                // サインインに必要なプリンシパルを作る
                // create claims
                List<Claim> claims = new List<Claim> { 
                                                        new Claim(ClaimTypes.Name, userViewModel.UserName),
                                                        new Claim("Password", loginModel.Password),
                                                        new Claim(ClaimTypes.Email, userViewModel.Email),
                                                        new Claim(ClaimTypes.StreetAddress, userViewModel.Adress),
                                                        new Claim(ClaimTypes.Role, userViewModel.RoleId.ToString()),
                                                     };

                var identity = new ClaimsIdentity(claims, CookieAuthenticationDefaults.AuthenticationScheme);
                var principal = new ClaimsPrincipal(identity);
                var authProperties = new Microsoft.AspNetCore.Authentication.AuthenticationProperties
                                                                  {
                                                                    AllowRefresh = false,
                                                                    ExpiresUtc = DateTimeOffset.Now.AddDays(1),
                                                                    IsPersistent = loginModel.RememberMe,
                                                                  };
                // 認証クッキーをレスポンスに追加
                // サインイン
                await HttpContext.SignInAsync(
                  CookieAuthenticationDefaults.AuthenticationScheme,
                  principal,
                  authProperties
                  );

                ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                //this.Session.SetString("UserName", loginModel.UserName);
                return RedirectToAction("Index", "Home");
            }
            else
            {
                ViewData.Add("errormess", "ユーザーIDまたはパスワードに誤りがあります");
                return View("Login");
            }
        }
        private ClaimsPrincipal ValidateToken(string jwtToken)
        {
            IdentityModelEventSource.ShowPII = true;

            SecurityToken validatedToken;
            TokenValidationParameters validationParameters = new TokenValidationParameters();

            validationParameters.ValidateLifetime = true;

            validationParameters.ValidAudience = _configuration["Tokens:Issuer"];
            validationParameters.ValidIssuer = _configuration["Tokens:Issuer"];
            validationParameters.IssuerSigningKey = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(_configuration["Tokens:Key"]));

            ClaimsPrincipal principal = new JwtSecurityTokenHandler().ValidateToken(jwtToken, validationParameters, out validatedToken);

            return principal;
        }
    }
}
