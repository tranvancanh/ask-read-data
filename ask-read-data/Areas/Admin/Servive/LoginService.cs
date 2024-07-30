using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using ask_read_data.Areas.Admin.Models;
using ask_read_data.Areas.Admin.Repository;
using ask_read_data.Areas.Dao;
using ask_read_data.Commons;

namespace read_data.Areas.Admin.Servive
{
    public class LoginService : ILoginViewModel
    {
        public bool UserLogin(LoginViewModel loginUser)
        {
            bool isLogin = false;
            int statusCode = -500;
            string passwordHash = Encryptor.MD5Hash(loginUser.Password);

            var ConnectionString = new GetConnectString().ConnectionString();
            using (var connection = new SqlConnection(ConnectionString))
            {
                try
                {
                    //Create the command object
                    SqlCommand cmd = new SqlCommand()
                    {
                        CommandText = LoginDao.SP_User_Login(),
                        Connection = connection,
                        CommandType = CommandType.Text
                    };
                    //パラメータ初期化
                    cmd.Parameters.Clear();
                    //Set SqlParameter
                    ///////////////////  SetParameter UserName  ////////////////////////////////////////////
                    SqlParameter UserName = new SqlParameter
                    {
                        ParameterName = "@UserName", //Parameter name defined in stored procedure
                        SqlDbType = SqlDbType.NVarChar, //Data Type of Parameter
                        Value = loginUser.UserName, //Value passes to the paramtere
                        Direction = ParameterDirection.Input //Specify the parameter as input
                    };
                    ///////////////////  SetParameter Password  ////////////////////////////////////////////
                    SqlParameter Password = new SqlParameter
                    {
                        ParameterName = "@Password", //Parameter name defined in stored procedure
                        SqlDbType = SqlDbType.NVarChar, //Data Type of Parameter
                        Value = loginUser.Password, //Value passes to the paramtere
                        Direction = ParameterDirection.Input //Specify the parameter as input
                    };
                    SqlParameter PasswordHash = new SqlParameter
                    {
                        ParameterName = "@PasswordHash", //Parameter name defined in stored procedure
                        SqlDbType = SqlDbType.NVarChar, //Data Type of Parameter
                        Value = passwordHash, //Value passes to the paramtere
                        Direction = ParameterDirection.Input //Specify the parameter as input
                    };
                    ///////////////////  SetParameter StausCode  ////////////////////////////////////////////
                    SqlParameter StausCode = new SqlParameter
                    {
                        ParameterName = "@StausCode", //Parameter name defined in stored procedure
                        SqlDbType = SqlDbType.Int, //Data Type of Parameter
                        Direction = ParameterDirection.Output //Specify the parameter as Output
                    };
                    cmd.Parameters.Add(UserName);
                    cmd.Parameters.Add(Password);
                    cmd.Parameters.Add(PasswordHash);
                    cmd.Parameters.Add(StausCode);

                    connection.Open();
                    cmd.ExecuteNonQuery();
                    statusCode = Convert.ToInt32(StausCode.Value);
                }
                catch
                {
                    throw;
                }
                finally
                {
                    //close connection
                    connection.Close();
                    // connection dispose
                    connection.Dispose();
                }
            }

            if (statusCode == 200)
            {
                isLogin = true;
            }
            else
            {
                isLogin = false;
            }

            return isLogin;
        }
    }
}
