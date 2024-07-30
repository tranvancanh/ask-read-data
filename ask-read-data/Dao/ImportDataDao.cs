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
    public class ImportDataDao
    {
        public static string SP_DataImportInsert()
        {
            var commandText = $@" INSERT INTO [DataImport]
                                   (
                                    WAYMD       ,
                                    SEQ         ,
                                    KATASIKI    ,
                                    MEISHO      ,
                                    FILLER1     ,
                                    OPT         ,
                                    JIKU        ,
                                    FILLER2     ,
                                    DAI         ,
                                    MC          ,
                                    SIMUKE      ,
                                    E0          ,
                                    BUBAN       ,
                                    TANTO       ,
                                    GR          ,
                                    KIGO        ,
                                    MAKR        ,
                                    KOSUU       ,
                                   KISYU       ,
                                   MEWISYO     ,
                                   FYMD        ,
                                   SEIHINCD    ,
                                   SEHINJNO    ,
                                   FileName    ,
                                   LineNumber  ,
                                   Position    ,
                                   CreateBy    ,
                                   CreateDateTime 
                                  ) 
                                VALUES 
                                    ( 
                                    @WAYMD       ,
                                    @SEQ         ,
                                    @KATASIKI    ,
                                    @MEISHO      ,
                                    @FILLER1     ,
                                    @OPT         ,
                                    @JIKU        ,
                                    @FILLER2     ,
                                    @DAI         ,
                                    @MC          ,
                                    @SIMUKE      ,
                                    @E0          ,
                                    @BUBAN       ,
                                    @TANTO       ,
                                    @GR          ,
                                    @KIGO        ,
                                    @MAKR        ,
                                    @KOSUU       ,
                                    @KISYU       ,
                                    @MEWISYO     ,
                                    @FYMD        ,
                                    @SEIHINCD    ,
                                    @SEHINJNO    ,
                                    @FileName    ,
                                    @LineNumber  ,
                                    @Position    ,
                                    @CreateBy    ,
                                    @CurrentDate 
                                    );";

            return commandText;
        }

        public static string SP_BU_Mastar_SelectInsertUpdateDelete()
        {
            var commandText = $@" DECLARE @CurrentDate datetime;  
                                  DECLARE @UpBy as nvarchar(50)
                                  DECLARE @CreateBy as nvarchar(50)
                                  DECLARE @UpDateTime as datetime
                                  DECLARE @CreateDateTime as datetime

                                  SET     @CurrentDate = FORMAT(GETDATE(), 'yyyy/MM/dd HH:mm:ss');
                                  SET     @UpBy = @User
                                  SET     @CreateBy = @User
                                  SET     @UpDateTime  = @CurrentDate
                                  SET     @CreateDateTime  = @CurrentDate
                                  SET     @StausCode = 0
  
                                ---------------------- Check User -----------------------------------------------
                                 IF @StatementType='Select'
                                 BEGIN
                                  ------------------------ Check Exits ------------------------------------------
                                  if (EXISTS( select * from [BU_Mastar] where BUBAN = @BUBAN ))
                                      Begin
                                        Set @StausCode = 200
                                        Return select @StausCode
                                      End
                                 else 
                                    Begin
                                        Set @StausCode = -500
                                        Return select @StausCode
                                    End
                                  END
                                  ---------------------- Update分  -----------------------------------------------
                                  ELSE IF @StatementType = 'Update'
                                  BEGIN
                                    ------------------------ Check Exits ------------------------------------------
                                  if (EXISTS( select * from [BU_Mastar] where BUBAN = @BUBAN ))
                                      Begin
                                       UPDATE [BU_Mastar] 
                                       SET MEWISYO = @MEWISYO,
                                           Nyusu = @Nyusu,
                                           UpDateTime = @CurrentDate,
                                           UpBy = @User

                                        WHERE BUBAN = @BUBAN 
                                         Set @StausCode = 200
                                         Return select @StausCode
                                       End
                                  else 
                                     Begin
                                         Set @StausCode = -500
                                         Return select @StausCode
                                     End
                                  END
                                  ----------------------------- Delete分 --------------------------------------------
                                  ELSE IF @StatementType = 'Delete'
                                  BEGIN
                                  ----------------------------- Check Exits ------------------------------------------
                                  if (EXISTS( select * from [BU_Mastar] where BUBAN = @BUBAN ))
                                       Begin
                                         DELETE FROM [BU_Mastar] 
    
                                        WHERE BUBAN = @BUBAN 
                                         Set @StausCode = 200
                                         Return select @StausCode
                                       End
                                  else 
                                     Begin
                                         Set @StausCode = -500
                                         Return select @StausCode
                                     End
       
                                  END
                                  ---------------------------- Insert分 --------------------------------------------
                                  ELSE IF @StatementType = 'Insert'
                                   BEGIN
                                  ---------------------------- Check Exits ------------------------------------------
                                   IF(NOT EXISTS( select * from [BU_Mastar] where BUBAN = @BUBAN ))
                                      Begin
                                         INSERT INTO [BU_Mastar] 
                                          (BUBAN   ,
                                           KIGO    ,
                                           MEWISYO ,
                                           KatakanaName,
                                           Nyusu   ,
                                           CreateDateTime,
                                           CreateBy)
                                         VALUES 
                                         (@BUBAN,
                                          @KIGO,
                                          @MEWISYO,
                                          @KatakanaName,
                                          @Nyusu, 
                                          @CreateDateTime, 
                                          @User)
                                          Set @StausCode = 200
                                         Return select @StausCode
                                       End
         
                                   END";

            return commandText;
        }

        public static string SP_File_Import_Log_Insert()
        {
            var commandText = $@" DECLARE @CurrentDate datetime;
                                 SET @StausCode = 0;
                                 SET @CurrentDate = FORMAT(GETDATE(), 'yyyy/MM/dd HH:mm:ss');
     
                                if (NOT EXISTS( select * from [File_Import_Log] where 1=1 
                                                            AND CreateDateTime = @CurrentDate 
                                                            AND FileName = @FileName
                                                            AND TotalLine = @TotalLine
                                                            AND MaxPosition = @MaxPosition
                                                            AND Createby = @User
                                
                                                            ))
                                   Begin
                                   INSERT INTO [File_Import_Log]
                                         (CreateDateTime, 
                                          FileName, 
                                          TotalLine, 
                                          MaxPosition,
                                          Createby)
                                   VALUES (@CurrentDate, 
                                          @FileName, 
                                          @TotalLine,
                                          @MaxPosition,
                                          @User)

                                     Set @StausCode = 200
                                     Return select @StausCode
                                    End
                             ";

            return commandText;
        }
        public static int DataImport_Before_CheckPosition(DateTime dateTime)
        {
            int position = 0;
            //connection
            SqlConnection connection = null;
            SqlCommand command = null;
            try
            {
                var ConnectionString = new GetConnectString().ConnectionString();
                using (connection = new SqlConnection(ConnectionString))
                {
                    //open
                    connection.Open();

                    //commmand
                    var commandText = $@"SELECT MAX(Position)  AS START_POSITION
                                        FROM [DataImport]
                                        WHERE (1=1)
                                        AND FORMAT(CreateDateTime, 'yyyy-MM-dd') = FORMAT(@dateTime, 'yyyy-MM-dd')";
                    command = new SqlCommand(commandText, connection);
                    command.Parameters.Add("@dateTime", System.Data.SqlDbType.DateTime).Value = dateTime;

                    var positionStart = command.ExecuteScalar();
                    if(Int32.TryParse(positionStart.ToString(), out int i))
                    {
                         if(i >= 0) { position = i; }
                         else { position = 0; }
                    }
                    else { position = 0; }
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                if(connection != null) { connection.Close(); connection.Dispose(); }
                if(command != null) { connection.Dispose(); }
            }
            return position;
        }

        public static Tuple<List<string>, List<string>> GetDropListItems(DateTime date)
        {
            var listItems = new List<string>();
            var listDateTimes = new List<string>();
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
                    var commandText = $@"SELECT TOP(5) FORMAT(CreateDateTime, 'yyyy-MM-dd') AS CreateDateTime
                                                  ,MAX(Position) AS MaxPosition
                                            FROM [DataImport]

                                            WHERE FORMAT(CreateDateTime, 'yyyy-MM-dd') <= FORMAT(CAST(@date AS date), 'yyyy-MM-dd')
                                            GROUP BY FORMAT(CreateDateTime, 'yyyy-MM-dd')

                                            ORDER BY CreateDateTime DESC , MaxPosition DESC";

                    SqlCommand command = new SqlCommand(commandText, connection);
                    command.Parameters.Clear();
                    command.Parameters.Add("@date", System.Data.SqlDbType.DateTime).Value = date;

                    reader = command.ExecuteReader();
                    var i = 1;
                    while (reader.Read())
                    {
                        var date1 = Convert.ToDateTime(reader["CreateDateTime"].ToString()).ToString("yyyy/MM/dd");
                        var maxPosition = Util.NullToBlank((object)reader["MaxPosition"]);
                        listItems.Add($@"前回{i} 取り込み日: " + date1 + " MaxPosition: " + maxPosition.ToString());
                        listDateTimes.Add(date1);
                        i++;
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
            var tuple = new Tuple<List<string>, List<string>>(listItems, listDateTimes);
            return tuple;
        }

        public static List<DataModel> GetDataOfLastTime(ImportViewModel viewModel)
        {
            var bubanType = viewModel.BubanType;
            var objList = new List<DataModel>();
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
                    switch (bubanType)
                    {
                        case "FL00R":
                        case "FRAME":
                            {
                                var createDateTime = viewModel.ItemValue;
                                if(connection.State != System.Data.ConnectionState.Open) { connection.Open(); }
                                var commandText = $@"SELECT TOP (2000) *
                                      FROM [DataImport]
                                      where (1=1)
                                      AND MEWISYO LIKE @bubanType
                                      AND FORMAT(CreateDateTime, 'yyyy-MM-dd') = FORMAT(CAST(@createDateTime AS date), 'yyyy-MM-dd')

                                      ORDER BY CreateDateTime DESC, Position DESC";
                                SqlCommand command = new SqlCommand(commandText, connection);
                                command.Parameters.Clear();
                                command.Parameters.Add("@bubanType", System.Data.SqlDbType.NVarChar).Value = '%' + bubanType + '%';
                                command.Parameters.Add("@createDateTime", System.Data.SqlDbType.DateTime).Value = Convert.ToDateTime(createDateTime);
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
                                var createDateTime = viewModel.ItemValue;
                                //commmand
                                connection = new SqlConnection(ConnectionString);
                                if (connection.State != System.Data.ConnectionState.Open) { connection.Open(); }
                                var commandText = $@"
                                                ------------------------------------------------------ Data -------------------------------------
                                                SELECT * FROM [DataImport]
                                                  WHERE (1=1)
                                                  AND FORMAT(CreateDateTime, 'yyyy-MM-dd') = FORMAT(CAST(@createDateTime AS date), 'yyyy-MM-dd')

                                                  ORDER BY CreateDateTime DESC, Position DESC";
                                SqlCommand command = new SqlCommand(commandText, connection);
                                command.Parameters.Clear();
                                command.Parameters.Add("@createDateTime", System.Data.SqlDbType.DateTime).Value = Convert.ToDateTime(createDateTime);

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
                if (connection != null) { connection.Close(); connection.Dispose(); }
                if (reader != null) { reader.Dispose(); }
            }

            return objList;
        }

        public static List<DataModel> GetDataOfLastTimeInit(DateTime date)
        {
            var objList = new List<DataModel>();
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
                    var commandText = $@"SELECT TOP(5) FORMAT(CreateDateTime, 'yyyy-MM-dd') AS CreateDateTime
                                                  ,MAX(Position) AS MaxPosition
                                            FROM [DataImport]

                                            WHERE FORMAT(CreateDateTime, 'yyyy-MM-dd') <= FORMAT(CAST(@date AS date), 'yyyy-MM-dd')
                                            GROUP BY FORMAT(CreateDateTime, 'yyyy-MM-dd')

                                            ORDER BY CreateDateTime DESC , MaxPosition DESC";

                    SqlCommand command = new SqlCommand(commandText, connection);
                    command.Parameters.Clear();
                    command.Parameters.Add("@date", System.Data.SqlDbType.DateTime).Value = date;

                    reader = command.ExecuteReader();
                    var date1 ="";
                    var maxPosition = 0;
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            date1 = Convert.ToDateTime(reader["CreateDateTime"].ToString()).ToString("yyyy/MM/dd");
                            maxPosition = Util.NullToBlank((object)reader["MaxPosition"]);
                            break;
                        }
                    }
                    else
                    {
                       return objList = new List<DataModel>();
                    }
                    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    //commmand
                    connection = new SqlConnection(ConnectionString);
                    if (connection.State != System.Data.ConnectionState.Open) { connection.Open(); }
                    commandText = $@"
                                                ------------------------------------------------------ Data -------------------------------------
                                                SELECT * FROM [DataImport]
                                                  WHERE (1=1)
                                                  AND FORMAT(CreateDateTime, 'yyyy-MM-dd') = FORMAT(CAST(@createDateTime AS date), 'yyyy-MM-dd')

                                                  ORDER BY CreateDateTime DESC, Position DESC";
                    command = new SqlCommand(commandText, connection);
                    command.Parameters.Clear();
                    command.Parameters.Add("@createDateTime", System.Data.SqlDbType.DateTime).Value = Convert.ToDateTime(date1);

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

        public static int DeleteDataOnToday(DateTime date)
        {
            var rowsAffected = 0;
            //connection
            SqlConnection connection = null;
            try
            {
                var ConnectionString = new GetConnectString().ConnectionString();
                using (connection = new SqlConnection(ConnectionString))
                {
                    //open
                    connection.Open();
                    var commandText = $@"
                                          if (EXISTS( select * from [DataImport] WHERE FORMAT([CreateDateTime], 'yyyy-MM-dd') = FORMAT(@date, 'yyyy-MM-dd') ))
                                            Begin
                                            DECLARE @datestart datetime
                                            SET @datestart = GETDATE()

                                             BEGIN TRY
                                               BEGIN TRANSACTION MYTRAN;
                                                 SET @datestart = (select FORMAT(CreateDateTime, 'yyyy-MM-dd HH:mm:ss')
                                                     FROM [DataImport]
                                                     WHERE FORMAT(CreateDateTime, 'yyyy-MM-dd') = FORMAT(@date, 'yyyy-MM-dd')
                                                     AND Position = 1);
                                                 PRINT @datestart
                                             ------------------------------------------------------ [dbo].[DataImport] ----------------------------------------------
                                                  DELETE FROM [DataImport] 
                                                            WHERE (1=1)
                                                            AND FORMAT(CreateDateTime, 'yyyy-MM-dd') = FORMAT(@date, 'yyyy-MM-dd');
                                             ------------------------------------------------------ [dbo].[File_Download_Log] ---------------------------------------
                                                 DELETE FROM [File_Download_Log] 
                                                 WHERE FORMAT([DownloadDateTime], 'yyyy-MM-dd') = FORMAT(@date, 'yyyy-MM-dd')
                                                          AND [DownloadDateTime] >= @datestart;

                                              ------------------------------------------------------ [dbo].[File_Import_Log] ----------------------------------------
                                                DELETE FROM [File_Import_Log]
                                                WHERE FORMAT([CreateDateTime], 'yyyy-MM-dd') = FORMAT(@date, 'yyyy-MM-dd');

                                                PRINT '>> COMMITING'
                                                COMMIT TRANSACTION MYTRAN;
                                              END TRY
                                            BEGIN CATCH
                                                PRINT '>> ROLLING BACK'
                                                ROLLBACK TRANSACTION MYTRAN; 
                                                THROW
                                            END CATCH
                                         End;";

                    SqlCommand command = new SqlCommand(commandText, connection);
                    command.Parameters.Clear();
                    command.Parameters.Add("@date", System.Data.SqlDbType.DateTime).Value = date;
                    rowsAffected = command.ExecuteNonQuery();
                  
                }
            }
            catch
            {
                throw;
            }
            finally
            {
                if (connection != null) { connection.Close(); connection.Dispose(); }
            }
            return rowsAffected;
        }
    }
}
