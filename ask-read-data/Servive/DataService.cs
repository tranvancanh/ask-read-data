using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using ask_read_data.Commons;
using ask_read_data.Models;
using ask_read_data.Repository;
using ask_tzn_funamiKD.Commons;

namespace ask_read_data.Servive
{
    public class DataService : IData
    {
        public List<DataModel> GetAll2000DataImport()
        {
            var objList = new List<DataModel>();
            var ConnectionString = new GetConnectString().ConnectionString;
            using (var connection = new SqlConnection(ConnectionString))
            {
                SqlDataReader reader = null;
                try
                {
                    connection.Open();
                    //Create the command object
                    SqlCommand cmd = new SqlCommand()
                                                        {
                                                            CommandText = "SP_BU_Mastar_GetAll2000Record",
                                                            Connection = connection,
                                                            CommandType = CommandType.StoredProcedure
                                                        };
                    cmd.Parameters.Clear();

                    //SQL実行
                    reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        var obj = new DataModel()
                        {
                            WAYMD = Convert.ToDateTime(reader["WAYMD"].ToString()),
                            SEQ = Util.NullToBlank((object)reader["SEQ"]),
                            KATASIKI = Util.NullToBlank(reader["KATASIKI"].ToString()),
                            MEISHO = Util.NullToBlank(reader["MEISHO"].ToString()),
                            JIKU = Util.NullToBlank(reader["JIKU"].ToString()),
                            BUBAN = Util.NullToBlank(reader["BUBAN"].ToString()),
                            KIGO = Util.NullToBlank(reader["KIGO"].ToString()),
                            MAKR = Util.NullToBlank(reader["MAKR"].ToString()),
                            MEWISYO = Util.NullToBlank(reader["MEWISYO"].ToString()),
                            FYMD = Convert.ToDateTime(reader["FYMD"].ToString())
                        };
                        objList.Add(obj);
                    }
                }
                catch
                {
                    throw;
                }
                finally
                {
                    if (connection != null)
                    {
                        connection.Close();
                    }
                    if (reader != null)
                    {
                        reader.Close();
                    }
                }
            }

            return objList;

        }

        public List<DataModel> SearchDataImport(DateTime importDate)
        {
            var objList = new List<DataModel>();
            var ConnectionString = new GetConnectString().ConnectionString;
            using (var connection = new SqlConnection(ConnectionString))
            {
                SqlDataReader reader = null;
                try
                {
                    connection.Open();
                    //Create the command object
                    SqlCommand cmd = new SqlCommand()
                                                        {
                                                            CommandText = "SP_BU_Mastar_SearchDataImport",
                                                            Connection = connection,
                                                            CommandType = CommandType.StoredProcedure
                                                        };
                    cmd.Parameters.Clear();
                    SqlParameter CreateDateTime = new SqlParameter
                                                            {
                                                                ParameterName = "@CreateDateTime",
                                                                SqlDbType = SqlDbType.DateTime,
                                                                Value = importDate,
                                                                Direction = ParameterDirection.Input
                                                            };
                    cmd.Parameters.Add(CreateDateTime);
                    //SQL実行
                    reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        var obj = new DataModel()
                        {
                            WAYMD = Convert.ToDateTime(reader["WAYMD"].ToString()),
                            SEQ = Util.NullToBlank((object)reader["SEQ"]),
                            KATASIKI = Util.NullToBlank(reader["KATASIKI"].ToString()),
                            MEISHO = Util.NullToBlank(reader["MEISHO"].ToString()),
                            JIKU = Util.NullToBlank(reader["JIKU"].ToString()),
                            BUBAN = Util.NullToBlank(reader["BUBAN"].ToString()),
                            KIGO = Util.NullToBlank(reader["KIGO"].ToString()),
                            MAKR = Util.NullToBlank(reader["MAKR"].ToString()),
                            MEWISYO = Util.NullToBlank(reader["MEWISYO"].ToString()),
                            FYMD = Convert.ToDateTime(reader["FYMD"].ToString())
                        };
                        objList.Add(obj);
                    }
                }
                catch
                {
                    throw;
                }
                finally
                {
                    if (connection != null)
                    {
                        connection.Close();
                    }
                    if (reader != null)
                    {
                        reader.Close();
                    }
                }
            }

            return objList;

        }
    }
}
