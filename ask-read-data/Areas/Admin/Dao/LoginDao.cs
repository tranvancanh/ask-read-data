using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ask_read_data.Areas.Dao
{
    public class LoginDao
    {
        public static string SP_User_Login()
        {
            var commandText = $@"set @StausCode = 0
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
                                  if (EXISTS( select * from [ask_datadb].[dbo].[LoginUser] where UserName = @UserName and Password = @PasswordHash))
                                    Begin
                                     Set @StausCode = 200
                                     Return select @StausCode
                                    End
                                  else 
                                Begin
                                    Set @StausCode = -500
                                    Return select @StausCode
                                End";

            return commandText;
        }


    }
}
