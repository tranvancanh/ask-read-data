using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ask_read_data.Areas.Admin.Dao
{
    public class UserDao
    {
        public static string SP_GetUser()
        {
            var commandText = $@"
                               SET @StausCode = 0
                               ---------------------- Check Data Legth ----------------------------------
                                if(LEN(@UserName) < 2  OR LEN(@UserName) > 50)
                                Begin
                                    Set @StausCode = -99
                                    Return select @StausCode
                                End
                                if(LEN(@Password) < 2  OR LEN(@Password) > 50)
                                Begin
                                   Set @StausCode = -99
                                   Return select @StausCode
                                End
                                ---------------------- Check Null --------------------------------------
                                if(@UserName is null OR @Password is null)
                                Begin
                                    Set @StausCode = -88
                                    Return select @StausCode
                                End
                               ---------------------- Check Exits -------------------------------------
                                if (EXISTS( select * from [Users] where UserName = @UserName and Password = @PasswordHash and IsActive = '1'))
                                 Begin
                                  select * from [Users] where UserName = @UserName and Password = @PasswordHash 
                                   Set @StausCode = 200
                                   Return select @StausCode
                                 End
                                else 
                               Begin
                                   Set @StausCode = -500
                                   Return select @StausCode
                               End

                                ";

            return commandText;
        }
    }
}
