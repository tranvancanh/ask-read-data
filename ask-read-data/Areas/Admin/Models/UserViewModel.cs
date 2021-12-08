using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace ask_read_data.Areas.Admin.Models
{
    public class UserViewModel
    {
        public int Id { get; set; }

        [Required(ErrorMessage = "ユーザーIDを入力してください")]
        [StringLength(50, ErrorMessage = "50文字以内で入力してください", MinimumLength = 2)]
        public string UserName { get; set; }

        [Required(ErrorMessage = "パスワードを入力してください")]
        [StringLength(50, ErrorMessage = "50文字以内で入力してください", MinimumLength = 2)]
        [DataType(DataType.Password)]
        public string Password { get; set; }

        [Required(ErrorMessage = "メールアドレスを入力してください")]
        [StringLength(50, ErrorMessage = "50文字以内で入力してください", MinimumLength = 2)]
        [DataType(DataType.EmailAddress)]
        public string Email { get; set; }

        [Required(ErrorMessage = "アドレスを入力してください")]
        [StringLength(50, ErrorMessage = "50文字以内で入力してください", MinimumLength = 2)]
        [DataType(DataType.Text)]
        public string Adress { get; set; }
        public int RoleId { get; set; }
        public bool IsActive { get; set; }
    }
}
