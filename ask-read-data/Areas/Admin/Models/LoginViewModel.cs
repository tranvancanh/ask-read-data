using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace ask_read_data.Areas.Admin.Models
{
    public class LoginViewModel
    {
        [Required(ErrorMessage = "ユーザーIDを入力してください")]
        [StringLength(50, ErrorMessage = "50文字以内で入力してください", MinimumLength=2)]
        public string UserName { get; set; }

        [Required(ErrorMessage = "パスワードを入力してください")]
        [StringLength(50, ErrorMessage = "50文字以内で入力してください", MinimumLength=2)]
        [DataType(DataType.Password)]
        public string Password { get; set; }

        [DefaultValue(false)]
        public bool RememberMe { get; set; }
    }
}
