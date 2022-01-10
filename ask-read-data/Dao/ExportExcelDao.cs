using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using ask_read_data.Commons;
using ask_read_data.Models;
using ask_tzn_funamiKD.Commons;

namespace ask_read_data.Dao
{
    public class ExportExcelDao
    {
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
                if (connection != null) { connection.Close(); }
                if(reader != null) { reader.Dispose(); }
            }
            var result = new Tuple<int, int>(Position, Renban);

            return result;
        }

        public static List<DataModel> GetRemainingDataOfLastTime(string bubanType)
        {
            var objList = new List<DataModel>();
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
                    switch (bubanType)
                    {
                        case "FL00R":
                        case "FRAME":
                            {
                                var lastDownloadDateTime = new DateTime();
                                var position = 0;
                                var renban = 0;
                                //commmand
                                var commandText = $@"SELECT TOP (10) [BubanMeiType]
                                      ,[LastDownloadDateTime]
                                      ,[Position]
                                      ,[ParetoRenban]
                                      ,[DownloadDateTime]
                                      ,[udownload]
                                  FROM [ask_datadb_test].[dbo].[File_Download_Log]

                                  where (1=1)
                                  AND BubanMeiType LIKE
                                  CASE WHEN @bubanType = 'FL00R' THEN '%FL00R ASSY%' 
                                       WHEN @bubanType = 'FRAME' THEN '%FRAME ASSY%' 
                                  ELSE
                                      '%ASSY%' 
                                  END
                                  AND FORMAT(LastDownloadDateTime, 'yyyy-MM-dd') < FORMAT(GETDATE(), 'yyyy-MM-dd')

                                  ORDER BY LastDownloadDateTime DESC, DownloadDateTime DESC";

                                SqlCommand command = new SqlCommand(commandText, connection);
                                command.Parameters.Clear();
                                command.Parameters.Add("@bubanType", System.Data.SqlDbType.NVarChar).Value = bubanType;

                                reader = command.ExecuteReader();
                                while (reader.Read())
                                {
                                    lastDownloadDateTime = Convert.ToDateTime((object)reader["LastDownloadDateTime"]);
                                    position = Util.NullToBlank((object)reader["Position"]);
                                    break;
                                }
                                /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                connection = new SqlConnection(ConnectionString);
                                if (connection.State == System.Data.ConnectionState.Closed) { connection.Open(); }
                                commandText = $@"SELECT TOP (100) *
                                      FROM [ask_datadb_test].[dbo].[DataImport]
                                      where (1=1)
                                      AND MEWISYO LIKE 
                                      CASE WHEN @BubanType = 'FL00R' THEN '%FL00R ASSY%' 
                                           WHEN @BubanType = 'FRAME' THEN '%FRAME ASSY%' 
                                      ELSE
                                          '%ASSY%'
                                      END
                                      AND FORMAT(CreateDateTime, 'yyyy-MM-dd') = FORMAT(CAST(@lastDownloadDateTime AS date), 'yyyy-MM-dd')
                                      AND Position > @position
                                      ORDER BY CreateDateTime DESC, Position DESC";
                                command = new SqlCommand(commandText, connection);
                                command.Parameters.Clear();
                                command.Parameters.Add("@bubanType", System.Data.SqlDbType.NVarChar).Value = bubanType;
                                command.Parameters.Add("@lastDownloadDateTime", System.Data.SqlDbType.DateTime).Value = lastDownloadDateTime;
                                command.Parameters.Add("@position", System.Data.SqlDbType.Int).Value = position;
                                reader = command.ExecuteReader();
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
                                        FYMD = Convert.ToDateTime(reader["FYMD"].ToString()),
                                        FileName = Util.NullToBlank(reader["FileName"].ToString()),
                                        Position = Util.NullToBlank((object)reader["Position"]),
                                        CreateDateTime = Convert.ToDateTime(reader["CreateDateTime"].ToString())
                                    };
                                    objList.Add(obj);
                                }
                                break;
                            }
                        case "ALL":
                            {
                                //commmand
                                var commandText = $@"SELECT * FROM (
                                                                   SELECT
                                                                   * FROM [ask_datadb_test].[dbo].[File_Download_Log]
                        
                                                                     WHERE (1=1)
                                                                     AND BubanMeiType LIKE '%FL00R ASSY%' 
                                                                     AND FORMAT(LastDownloadDateTime, 'yyyy-MM-dd') < FORMAT(GETDATE(), 'yyyy-MM-dd')
                                                                     
                                                                     ORDER BY LastDownloadDateTime DESC, DownloadDateTime DESC
                                                                     OFFSET 0 ROWS FETCH NEXT 1 ROWS ONLY

                                UNION ALL

                                                                   SELECT
                                                                     * FROM [ask_datadb_test].[dbo].[File_Download_Log]
                                                                     
                                                                       WHERE (1=1)
                                                                       AND BubanMeiType LIKE '%FRAME ASSY%' 
                                                                       AND FORMAT(LastDownloadDateTime, 'yyyy-MM-dd') < FORMAT(GETDATE(), 'yyyy-MM-dd')
                                                                       
                                                                       ORDER BY LastDownloadDateTime DESC, DownloadDateTime DESC
                                                                       OFFSET 0 ROWS FETCH NEXT 1 ROWS ONLY
                                                                       ) NEWTABLE";
                                connection = new SqlConnection(ConnectionString);
                                if (connection.State == System.Data.ConnectionState.Closed) { connection.Open(); }
                                SqlCommand command = new SqlCommand(commandText, connection);
                                command.Parameters.Clear();
                                reader = command.ExecuteReader();
                                var BubanMeiType1 = "";
                                var lastDownloadDateTime1 = new DateTime(1900, 01, 01, 00, 00, 00);
                                var position1 = 0;

                                var BubanMeiType2 = "";
                                var lastDownloadDateTime2 = new DateTime(1900, 01, 01, 00, 00, 00);
                                var position2 = 0;
                                while (reader.Read())
                                {
                                    if(Util.NullToBlank(reader["BubanMeiType"].ToString()).Contains("FL00R"))
                                    {
                                        BubanMeiType1 = "FL00R ASSY";
                                        lastDownloadDateTime1 = Convert.ToDateTime((object)reader["LastDownloadDateTime"]);
                                        position1 = Util.NullToBlank((object)reader["Position"]);
                                    }
                                    else if(Util.NullToBlank(reader["BubanMeiType"].ToString()).Contains("FRAME"))
                                    {
                                        BubanMeiType2 = "FRAME ASSY";
                                        lastDownloadDateTime2 = Convert.ToDateTime((object)reader["LastDownloadDateTime"]);
                                        position2 = Util.NullToBlank((object)reader["Position"]);
                                    }
                                }
                                /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                connection = new SqlConnection(ConnectionString);
                                if (connection.State == System.Data.ConnectionState.Closed) { connection.Open(); }
                                commandText = $@"
                                                ------------------------------------------------------ FL00R Data -------------------------------------
                                                SELECT * FROM [ask_datadb_test].[dbo].[DataImport]
                                                  WHERE (1=1)
                                                  AND MEWISYO LIKE '%FL00R%' 
                                                  AND FORMAT(CreateDateTime, 'yyyy-MM-dd') = FORMAT(CAST(@lastDownloadDateTime1 AS date), 'yyyy-MM-dd')  
                                                  AND [Position] > @position1

                                                UNION ALL 
                                                ------------------------------------------------------ FRAME Data --------------------------------------
                                                SELECT * FROM [ask_datadb_test].[dbo].[DataImport]
                                                  WHERE (1=1)
                                                  AND MEWISYO LIKE '%FRAME%' 
                                                  AND FORMAT(CreateDateTime, 'yyyy-MM-dd') = FORMAT(CAST(@lastDownloadDateTime2 AS date), 'yyyy-MM-dd')  
                                                  AND [Position] > @position2
                                                  ORDER BY MEWISYO ASC, CreateDateTime DESC, Position DESC";
                                command = new SqlCommand(commandText, connection);
                                command.Parameters.Clear();
                                command.Parameters.Add("@lastDownloadDateTime1", System.Data.SqlDbType.DateTime).Value = lastDownloadDateTime1;
                                command.Parameters.Add("@position1", System.Data.SqlDbType.Int).Value = position1;
                                //--------------------------------------------------------------------------------------------------------------------------
                                command.Parameters.Add("@lastDownloadDateTime2", System.Data.SqlDbType.DateTime).Value = lastDownloadDateTime2;
                                command.Parameters.Add("@position2", System.Data.SqlDbType.Int).Value = position2;

                                reader = command.ExecuteReader();
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
                                        FYMD = Convert.ToDateTime(reader["FYMD"].ToString()),
                                        FileName = Util.NullToBlank(reader["FileName"].ToString()),
                                        Position = Util.NullToBlank((object)reader["Position"]),
                                        CreateDateTime = Convert.ToDateTime(reader["CreateDateTime"].ToString())
                                    };
                                    objList.Add(obj);
                                }
                                break;
                            }
                    }
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                if (connection != null) { connection.Close(); }
                if (reader != null) { reader.Dispose(); }
            }

            return objList;
        }
    }
}
