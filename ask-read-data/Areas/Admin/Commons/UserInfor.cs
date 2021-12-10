using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using ask_read_data.Areas.Admin.Models;
using Microsoft.AspNetCore.Mvc;

namespace ask_read_data.Areas.Admin.Commons
{
    [Area(areaName: "Admin")]
    public class UserInfor : Microsoft.AspNetCore.Mvc.Controller
    {
        public UserViewModel UserInfo()
        {
            //   Get Current User Claims   /

            var user1 = User.Claims.Where(c => c.Type == ClaimTypes.Name).First().Value;
            var claims1 = User.Claims.ToList();
            var user2 = User.Claims.Where(c => c.Type == ClaimTypes.Name).First().Value;
            // get user claims  
            var claims = User.Claims.ToList();
            var user = new UserViewModel();
            user.UserName = Convert.ToString(claims.Where(x => x.Type == ClaimTypes.Name).First().Value);
            user.Email = Convert.ToString(claims.Where(x => x.Type == ClaimTypes.Email).First().Value);
            user.Adress = Convert.ToString(claims.Where(x => x.Type == ClaimTypes.StreetAddress).First().Value);
            user.RoleId = Convert.ToInt32(claims.Where(x => x.Type == ClaimTypes.Role).First().Value);

            return user;
        }
        //public string GetUserId(this ClaimsPrincipal principal)
        //{
        //    if (principal == null)
        //        throw new ArgumentNullException(nameof(principal));

        //    return principal.FindFirst(ClaimTypes.NameIdentifier)?.Value;
        //}
    }
}
