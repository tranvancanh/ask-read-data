using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using ask_read_data.Commons;
using ask_read_data.Models;
using ask_read_data.Models.Entity;
using ask_read_data.Models.ViewModel;
using ask_tzn_funamiKD.Commons;

namespace ask_read_data.Dao
{
    public class ExportExcelDao
    {
        public static bool SP_DataImport_CheckData(DateTime CreateDateTime, string bubanType)
        {
            var isCheck = false;
            //connection
            SqlConnection connection = null;
            SqlDataReader reader = null;
            try
            {
                var ConnectionString = new GetConnectString().ConnectionString;
                using (connection = new SqlConnection(ConnectionString))
                {
                    //open
                    connection.Open();

                    //commmand
                    var commandText = $@"select * from [ask_datadb_test].[dbo].[DataImport]
                                           WHERE (1=1)
	                                            AND FORMAT(CreateDateTime, 'yyyy-MM-dd') = FORMAT(@CreateDateTime, 'yyyy-MM-dd') 
	                                            AND MEWISYO LIKE '%'+ @BubanType +'%'";

                    SqlCommand command = new SqlCommand(commandText, connection);
                    command.Parameters.Clear();
                    command.Parameters.Add("@CreateDateTime", System.Data.SqlDbType.DateTime).Value = CreateDateTime;
                    command.Parameters.Add("@bubanType", System.Data.SqlDbType.NVarChar).Value = '%' + bubanType + '%';

                    reader = command.ExecuteReader();
                    if (reader.HasRows)
                    {
                        isCheck = true;
                    }
                     else { isCheck = false; }
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                if (connection != null) { connection.Close(); connection.Dispose(); }
                if (reader != null) { reader.Dispose(); }
            }

            return isCheck;
        }
        public static Tuple<int, int> GetPositonParetoRenban_Before_Download(DateTime date, string bubanType)
        {
            var Position = 0;
            var Renban = 0;
            //connection
            SqlConnection connection = null;
            SqlDataReader reader = null;
            try
            {
                var ConnectionString = new GetConnectString().ConnectionString;
                using (connection = new SqlConnection(ConnectionString))
                {
                    //open
                    connection.Open();

                    //commmand
                    var commandText = $@"SELECT TOP (10) 
                                       [BubanMeiType]
                                      ,[LastDownloadDateTime]
                                      ,[Position]
                                      ,[ParetoRenban]
                                      ,[DownloadDateTime]
                                  FROM [ask_datadb_test].[dbo].[File_Download_Log]
                                  WHERE FORMAT(LastDownloadDateTime, 'yyyy-MM-dd') = FORMAT(@date, 'yyyy-MM-dd') 
                                  AND BubanMeiType like @bubanType
                                  ORDER BY LastDownloadDateTime DESC, DownloadDateTime DESC";

                    SqlCommand command = new SqlCommand(commandText, connection);
                    command.Parameters.Clear();
                    command.Parameters.Add("@date", System.Data.SqlDbType.DateTime).Value = date;
                    command.Parameters.Add("@bubanType", System.Data.SqlDbType.NVarChar).Value = '%' + bubanType + '%';

                    reader = command.ExecuteReader();
                    while (reader.Read())
                    {
                        Position = Util.NullToBlank((object)reader["Position"]);
                        Renban = Util.NullToBlank((object)reader["ParetoRenban"]);
                        break;
                    }

                }
            }
            catch
            {
                throw;
            }
            finally
            {
                if (connection != null) { connection.Close(); connection.Dispose(); }
                if(reader != null) { reader.Dispose(); }
            }
            var result = new Tuple<int, int>(Position, Renban);

            return result;
        }

        public static List<FileDownloadLogModel> GetDownloadHistory(DateTime date)
        {
            var objList = new List<FileDownloadLogModel>();
            //connection
            SqlConnection connection = null;
            SqlDataReader reader = null;
            try
            {
                var ConnectionString = new GetConnectString().ConnectionString;
                using (connection = new SqlConnection(ConnectionString))
                {
                    //open
                    connection.Open();
                    //commmand
                    var commandText = $@"SELECT * FROM (
                                        ------------------------------------------------------ FL00R Data -------------------------------------
                                                        SELECT
                                                        * FROM [ask_datadb_test].[dbo].[File_Download_Log]
                        
                                                            WHERE (1=1)
                                                            AND BubanMeiType LIKE '%FL00R ASSY%' 
                                                            AND FORMAT(LastDownloadDateTime, 'yyyy-MM-dd') <= FORMAT(@date, 'yyyy-MM-dd')
                        
                                                            ORDER BY LastDownloadDateTime DESC, DownloadDateTime DESC
                                                            OFFSET 0 ROWS FETCH NEXT 5 ROWS ONLY
                                        ------------------------------------------------------ FRAME Data -------------------------------------
                                                    UNION ALL
                                                        SELECT
                                                            * FROM [ask_datadb_test].[dbo].[File_Download_Log]
                                
                                                            WHERE (1=1)
                                                            AND BubanMeiType LIKE '%FRAME ASSY%' 
                                                            AND FORMAT(LastDownloadDateTime, 'yyyy-MM-dd') <= FORMAT(@date, 'yyyy-MM-dd')
                                
                                                            ORDER BY LastDownloadDateTime DESC, DownloadDateTime DESC
                                                            OFFSET 0 ROWS FETCH NEXT 5 ROWS ONLY
                                                      ) NEWTABLE";

                    SqlCommand command = new SqlCommand(commandText, connection);
                    command.Parameters.Clear();
                    command.Parameters.Add("@date", System.Data.SqlDbType.DateTime).Value = date;
                    reader = command.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            var obj = new FileDownloadLogModel()
                            {
                                BubanMeiType = Util.NullToBlank(reader["BubanMeiType"].ToString()),
                                LastDownloadDateTime = Convert.ToDateTime(reader["LastDownloadDateTime"].ToString()),
                                Position = Util.NullToBlank((object)reader["Position"]),
                                ParetoRenban = Util.NullToBlank((object)reader["ParetoRenban"]),
                                DownloadDateTime = Convert.ToDateTime(reader["DownloadDateTime"].ToString())
                            };
                           objList.Add(obj);
                        }
                    }
                     else { return objList = new List<FileDownloadLogModel>(); }
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                if (connection != null) { connection.Close(); connection.Dispose(); }
                if (reader != null) { reader.Dispose(); }
            }

            return objList;
        }
        public static Tuple<DateTime, int, int> GetPositionParetoRenbanLasttime(string bubanType)
        {
            var DateTime = new DateTime();
            var Position = 0;
            var Renban = 0;
            //connection
            SqlConnection connection = null;
            SqlDataReader reader = null;
            try
            {
                var ConnectionString = new GetConnectString().ConnectionString;
                using (connection = new SqlConnection(ConnectionString))
                {
                    //open
                    connection.Open();

                    //commmand
                    var commandText = $@"SELECT TOP (10) * 
                                                      FROM [ask_datadb_test].[dbo].[File_Download_Log]
                                                      WHERE (1=1)
                                                      AND BubanMeiType like @bubanType
                                                     ORDER BY LastDownloadDateTime DESC, DownloadDateTime DESC";

                    SqlCommand command = new SqlCommand(commandText, connection);
                    command.Parameters.Clear();
                    command.Parameters.Add("@bubanType", System.Data.SqlDbType.NVarChar).Value = '%' + bubanType + '%';
                    reader = command.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            DateTime = Convert.ToDateTime(reader["LastDownloadDateTime"].ToString());
                            Position = Util.NullToBlank((object)reader["Position"]);
                            Renban = Util.NullToBlank((object)reader["ParetoRenban"]);
                            break;
                        }
                    }
                    else
                    {
                        DateTime = new DateTime();
                        Position = 0;
                        Renban = 0;
                    }
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                if (connection != null) { connection.Close(); connection.Dispose(); }
                if (reader != null) { reader.Dispose(); }
            }
            var result = new Tuple<DateTime, int, int>(DateTime, Position, Renban);

            return result;
        }
    }
}
