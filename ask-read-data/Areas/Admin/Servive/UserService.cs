using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using ask_read_data.Areas.Admin.Dao;
using ask_read_data.Areas.Admin.Models;
using ask_read_data.Areas.Admin.Repository;
using ask_read_data.Commons;
using ask_tzn_funamiKD.Commons;

namespace ask_read_data.Areas.Admin.Servive
{
    public class UserService : IUser
    {
        public UserViewModel GetUser(string userName, string password)
        {
            string passwordHash = Encryptor.MD5Hash(password);
            int statusCode = -500;
            var obj = new UserViewModel();
         
                var ConnectionString = new GetConnectString().ConnectionString();
                using (var connection = new SqlConnection(ConnectionString))
                {
                try
                {
                    //Create the command object
                    SqlCommand cmd = new SqlCommand()
                    {
                        CommandText = UserDao.SP_GetUser(),
                        Connection = connection,
                        CommandType = CommandType.Text
                    };
                    //パラメータ初期化
                    cmd.Parameters.Clear();
                    //Set SqlParameter
                    ///////////////////  SetParameter UserName  ////////////////////////////////////////////
                    SqlParameter UserName = new SqlParameter
                    {
                        ParameterName = "@UserName",
                        SqlDbType = SqlDbType.NVarChar,
                        Value = userName,
                        Direction = ParameterDirection.Input
                    };
                    ///////////////////  SetParameter PasswordHash  ////////////////////////////////////////////
                    SqlParameter Password = new SqlParameter
                    {
                        ParameterName = "@Password",
                        SqlDbType = SqlDbType.NVarChar,
                        Value = password,
                        Direction = ParameterDirection.Input
                    };
                    ///////////////////  SetParameter PasswordHash  ////////////////////////////////////////////
                    SqlParameter PasswordHash = new SqlParameter
                    {
                        ParameterName = "@PasswordHash",
                        SqlDbType = SqlDbType.NVarChar,
                        Value = passwordHash,
                        Direction = ParameterDirection.Input
                    };

                    ///////////////////  SetParameter StausCode  ////////////////////////////////////////////
                    SqlParameter StausCode = new SqlParameter
                    {
                        ParameterName = "@StausCode",
                        SqlDbType = SqlDbType.Int,
                        Direction = ParameterDirection.Output
                    };
                    cmd.Parameters.Add(UserName);
                    cmd.Parameters.Add(Password);
                    cmd.Parameters.Add(PasswordHash);
                    cmd.Parameters.Add(StausCode);

                    connection.Open();
                    //SQL実行
                    SqlDataReader reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        obj = new UserViewModel()
                        {
                            UserName = Util.NullToBlank(reader["UserName"].ToString()),
                            Password = Util.NullToBlank(reader["Password"].ToString()),
                            Email = Util.NullToBlank(reader["Email"].ToString()),
                            Adress = Util.NullToBlank(reader["Adress"].ToString()),
                            RoleId = Util.NullToBlank(reader["RoleId"])
                        };
                    }
                    reader.Close();
                    statusCode = Convert.ToInt32(StausCode.Value);
                    if (statusCode != 200 || obj == new UserViewModel())
                    {
                        return obj = null;
                    }
                }
                catch
                {
                    throw;
                }
                finally
                {
                    connection.Close();
                }
            }
            
            return obj;
        }
    }
}
