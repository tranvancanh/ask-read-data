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
        public static string SP_File_Download_Log_Insert()
        {
            var commandText = $@" DECLARE @CurrentDate datetime;
                                 SET @StausCode = 0
                                 SET @CurrentDate = FORMAT(GETDATE(), 'yyyy/MM/dd HH:mm:ss');
 
                              if (NOT EXISTS( select * from [File_Download_Log] where 1=1 
                                                AND BubanMeiType = @BubanMeiType 
                                                AND LastDownloadDateTime = @LastDownloadDateTime
                                                AND DownloadDateTime = @CurrentDate
                                                AND udownload = @User))
                                   Begin
                                   INSERT INTO [File_Download_Log] 
                                          (BubanMeiType, 
                                           LastDownloadDateTime, 
                                           Position, 
                                           ParetoRenban, 
                                           DownloadDateTime,
                                           udownload)
                                   VALUES (@BubanMeiType, 
                                           @LastDownloadDateTime, 
                                           @Position,
                                           @ParetoRenban,
                                           @CurrentDate,
                                           @User)

                                    Set @StausCode = 200
                                  End
                             else 
                                Begin
                                    Set @StausCode = -500
                                End";

            return commandText;
        }
        public static string SP_GetLastDowloadInfo()
        {
            var commandText = $@" SELECT TOP (100) 
                                   [BubanMeiType]
                                  ,[LastDownloadDateTime]
                                  ,[Position]
                                  ,[ParetoRenban]
                                  ,[DownloadDateTime]
                                  ,[udownload]
                              FROM [File_Download_Log]

                                WHERE (1=1)
                                AND LastDownloadDateTime <= @DateTime
                                AND BubanMeiType = @BubanMeiType

                                 ORDER BY LastDownloadDateTime DESC, DownloadDateTime DESC";
            return commandText;
        }
        public static string SP_DataImport_Flame_AssyAndFloor_Assy()
        {
            var commandText = $@"
                                  ------------------------------  開始日の残りデータをチェック　--------------------------------------
                                      SELECT 
                                         [WAYMD]
                                        ,[SEQ]
                                        ,[KATASIKI]
                                        ,[MEISHO]
                                        ,[FILLER1]
                                        ,[OPT]
                                        ,[JIKU]
                                        ,[FILLER2]
                                        ,[DAI]
                                        ,[MC]
                                        ,[SIMUKE]
                                        ,[E0]
                                        ,[BUBAN]
                                        ,[TANTO]
                                        ,[GR]
                                        ,[KIGO]
                                        ,[MAKR]
                                        ,[KOSUU]
                                        ,[KISYU]
                                        ,[MEWISYO]
                                        ,[FYMD]
                                        ,[SEIHINCD]
                                        ,[SEHINJNO]
                                        ,[FileName]
                                        ,[LineNumber]
                                        ,[Position]
                                        ,[CreateDateTime]
                                        ,FORMAT(CreateDateTime, 'yyyy-MM-dd') as CreateDateTimeFormat
        
        
                                        FROM [DataImport] as a
                                        WHERE (1=1) 
                                        AND FORMAT(CreateDateTime, 'yyyy-MM-dd') = FORMAT(@StartDateTime, 'yyyy-MM-dd')
                                        AND MEWISYO LIKE '%'+ @BubanType +'%'
                                        AND Position > @StartPosition
        
                                       UNION ALL
                                      --------------------------------------- 今までのデータをチェック　-----------------------------------------------
                                          SELECT 
                                             [WAYMD]
                                            ,[SEQ]
                                            ,[KATASIKI]
                                            ,[MEISHO]
                                            ,[FILLER1]
                                            ,[OPT]
                                            ,[JIKU]
                                            ,[FILLER2]
                                            ,[DAI]
                                            ,[MC]
                                            ,[SIMUKE]
                                            ,[E0]
                                            ,[BUBAN]
                                            ,[TANTO]
                                            ,[GR]
                                            ,[KIGO]
                                            ,[MAKR]
                                            ,[KOSUU]
                                            ,[KISYU]
                                            ,[MEWISYO]
                                            ,[FYMD]
                                            ,[SEIHINCD]
                                            ,[SEHINJNO]
                                            ,[FileName]
                                            ,[LineNumber]
                                            ,[Position]
                                            ,[CreateDateTime]
                                            ,FORMAT(CreateDateTime, 'yyyy-MM-dd') as CreateDateTimeFormat
        
        
                                        FROM [DataImport]
                                        WHERE (1=1)
                                        AND FORMAT(CreateDateTime, 'yyyy-MM-dd') > FORMAT(@StartDateTime, 'yyyy-MM-dd')
                                        AND FORMAT(CreateDateTime, 'yyyy-MM-dd') <= FORMAT(GETDATE(), 'yyyy-MM-dd')
                                        AND MEWISYO LIKE '%'+ @BubanType +'%'
        
                                       ORDER BY [CreateDateTime] ASC, [Position] ASC ";

            return commandText;
        }
        public static bool SP_DataImport_CheckData(DateTime CreateDateTime, string bubanType)
        {
            var isCheck = false;
            //connection
            SqlConnection connection = null;
            SqlDataReader reader = null;
            try
            {
                var ConnectionString = new GetConnectString().ConnectionString();
                using (connection = new SqlConnection(ConnectionString))
                {
                    //open
                    connection.Open();

                    //commmand
                    var commandText = $@"select * from [DataImport]
                                           WHERE (1=1)
	                                            AND FORMAT(CreateDateTime, 'yyyy-MM-dd') = FORMAT(@CreateDateTime, 'yyyy-MM-dd') 
	                                            AND MEWISYO LIKE @BubanType";

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
                var ConnectionString = new GetConnectString().ConnectionString();
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
                                  FROM [File_Download_Log]
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
                var ConnectionString = new GetConnectString().ConnectionString();
                using (connection = new SqlConnection(ConnectionString))
                {
                    //open
                    connection.Open();
                    //commmand
                    var commandText = $@"SELECT * FROM (
                                        ------------------------------------------------------ FL00R Data -------------------------------------
                                                        SELECT
                                                        * FROM [File_Download_Log]
                        
                                                            WHERE (1=1)
                                                            AND BubanMeiType LIKE '%FL00R ASSY%' 
                                                            AND FORMAT(LastDownloadDateTime, 'yyyy-MM-dd') <= FORMAT(@date, 'yyyy-MM-dd')
                        
                                                            ORDER BY LastDownloadDateTime DESC, DownloadDateTime DESC
                                                            OFFSET 0 ROWS FETCH NEXT 5 ROWS ONLY
                                        ------------------------------------------------------ FRAME Data -------------------------------------
                                                    UNION ALL
                                                        SELECT
                                                            * FROM [File_Download_Log]
                                
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
                var ConnectionString = new GetConnectString().ConnectionString();
                using (connection = new SqlConnection(ConnectionString))
                {
                    //open
                    connection.Open();

                    //commmand
                    var commandText = $@"SELECT TOP (10) * 
                                                      FROM [File_Download_Log]
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
