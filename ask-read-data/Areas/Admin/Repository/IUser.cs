using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ask_read_data.Areas.Admin.Models;

namespace ask_read_data.Areas.Admin.Repository
{
    public interface IUser
    {
        UserViewModel GetUser(string UserName, string Password);
    }
}
