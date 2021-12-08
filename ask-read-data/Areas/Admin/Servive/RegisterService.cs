using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using ask_read_data.Areas.Admin.Models;
using ask_read_data.Areas.Admin.Repository;
using ask_read_data.Commons;
using ask_tzn_funamiKD.Commons;

namespace ask_read_data.Areas.Admin.Servive
{
    public class RegisterService : IRegisterService
    {
        private string statementType = "";
        public UserViewModel GetListUser(UserViewModel userModel)
        {
            statementType = "Select";
            int statusCode = -500;
            var obj = new UserViewModel();

            var ConnectionString = new GetConnectString().ConnectionString;
            using (var connection = new SqlConnection(ConnectionString))
            {
                try
                {
                    //Create the command object
                    SqlCommand cmd = new SqlCommand()
                    {
                        CommandText = "SP_Users",
                        Connection = connection,
                        CommandType = CommandType.StoredProcedure
                    };
                    //パラメータ初期化
                    cmd.Parameters.Clear();
                    //Set SqlParameter
                    ///////////////////  SetParameter UserName  ////////////////////////////////////////////
                    SqlParameter UserName = new SqlParameter
                    {
                        ParameterName = "@UserName",
                        SqlDbType = SqlDbType.NVarChar,
                        Value = userModel.UserName,
                        Direction = ParameterDirection.Input
                    };
                    ///////////////////  SetParameter Password  ////////////////////////////////////////////
                    SqlParameter Password = new SqlParameter
                    {
                        ParameterName = "@Password",
                        SqlDbType = SqlDbType.NVarChar,
                        Value = userModel.Password,
                        Direction = ParameterDirection.Input
                    };
                    ///////////////////  SetParameter Email  ////////////////////////////////////////////
                    SqlParameter Email = new SqlParameter
                    {
                        ParameterName = "@Email",
                        SqlDbType = SqlDbType.NVarChar,
                        Value = userModel.Email,
                        Direction = ParameterDirection.Input
                    };
                    ///////////////////  SetParameter Adress  ////////////////////////////////////////////
                    SqlParameter Adress = new SqlParameter
                    {
                        ParameterName = "@Adress",
                        SqlDbType = SqlDbType.NVarChar,
                        Value = userModel.Adress,
                        Direction = ParameterDirection.Input
                    };
                    ///////////////////  SetParameter StausCode  ////////////////////////////////////////////
                    SqlParameter StausCode = new SqlParameter
                    {
                        ParameterName = "@StausCode",
                        SqlDbType = SqlDbType.Int,
                        Direction = ParameterDirection.Output
                    };
                    ///////////////////  SetParameter StatementType  ////////////////////////////////////////////
                    SqlParameter StatementType = new SqlParameter
                    {
                        ParameterName = "@StatementType",
                        SqlDbType = SqlDbType.NVarChar,
                        Value = statementType,
                        Direction = ParameterDirection.Input
                    };
                    cmd.Parameters.Add(UserName);
                    cmd.Parameters.Add(Password);
                    cmd.Parameters.Add(Email);
                    cmd.Parameters.Add(Adress);
                    cmd.Parameters.Add(StausCode);
                    cmd.Parameters.Add(StatementType);


                    connection.Open();
                    //SQL実行
                    SqlDataReader reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        obj = new UserViewModel()
                        {
                            UserName = Util.GetDataDB(reader, "UserName"),
                            Password = Util.GetDataDB(reader, "Password"),
                            Email = Util.GetDataDB(reader, "Email"),
                            Adress = Util.GetDataDB(reader, "Adress"),
                            RoleId = Convert.ToInt32(Util.GetDataDB(reader, "RoleId"))
                        };
                    }
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

            return obj;
        }

        public UserViewModel UpdateUser(UserViewModel userModel)
        {
            statementType = "Update";
            throw new NotImplementedException();
        }
        public bool DeleteUser(UserViewModel userModel)
        {
            statementType = "Delete";
            throw new NotImplementedException();
        }


    }
}
