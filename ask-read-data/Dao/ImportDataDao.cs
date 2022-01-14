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
        public static int DataImport_Before_CheckPosition(DateTime dateTime)
        {
            int position = 0;
            //connection
            SqlConnection connection = null;
            SqlCommand command = null;
            try
            {
                var ConnectionString = new GetConnectString().ConnectionString;
                using (connection = new SqlConnection(ConnectionString))
                {
                    //open
                    connection.Open();

                    //commmand
                    var commandText = $@"SELECT MAX(Position)  AS START_POSITION
                                        FROM [ask_datadb_test].[dbo].[DataImport]
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
                var ConnectionString = new GetConnectString().ConnectionString;
                using (connection = new SqlConnection(ConnectionString))
                {
                    //open
                    connection.Open();
                    var commandText = $@"SELECT TOP(5) FORMAT(CreateDateTime, 'yyyy-MM-dd') AS CreateDateTime
                                                  ,MAX(Position) AS MaxPosition
                                            FROM [ask_datadb_test].[dbo].[DataImport]

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
                        listItems.Add($@"前回{i} 作成時間: " + date1 + " MaxPosition: " + maxPosition.ToString());
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
                                var createDateTime = viewModel.ItemValue;
                                if(connection.State != System.Data.ConnectionState.Open) { connection.Open(); }
                                var commandText = $@"SELECT TOP (2000) *
                                      FROM [ask_datadb_test].[dbo].[DataImport]
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
                                                SELECT * FROM [ask_datadb_test].[dbo].[DataImport]
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
                var ConnectionString = new GetConnectString().ConnectionString;
                using (connection = new SqlConnection(ConnectionString))
                {
                    //open
                    connection.Open();
                    var commandText = $@"SELECT TOP(5) FORMAT(CreateDateTime, 'yyyy-MM-dd') AS CreateDateTime
                                                  ,MAX(Position) AS MaxPosition
                                            FROM [ask_datadb_test].[dbo].[DataImport]

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
                                                SELECT * FROM [ask_datadb_test].[dbo].[DataImport]
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
                var ConnectionString = new GetConnectString().ConnectionString;
                using (connection = new SqlConnection(ConnectionString))
                {
                    //open
                    connection.Open();
                    var commandText = $@"
                                          if (EXISTS( select * from [ask_datadb_test].[dbo].[DataImport] WHERE FORMAT([CreateDateTime], 'yyyy-MM-dd') = FORMAT(@date, 'yyyy-MM-dd') ))
                                            Begin
                                            DECLARE @datestart datetime
                                            SET @datestart = GETDATE()

                                             BEGIN TRY
                                               BEGIN TRANSACTION MYTRAN;
                                                 SET @datestart = (select FORMAT(CreateDateTime, 'yyyy-MM-dd HH:mm:ss')
                                                     FROM [ask_datadb_test].[dbo].[DataImport]
                                                     WHERE FORMAT(CreateDateTime, 'yyyy-MM-dd') = FORMAT(@date, 'yyyy-MM-dd')
                                                     AND Position = 1);
                                                 PRINT @datestart
                                             ------------------------------------------------------ [dbo].[DataImport] ----------------------------------------------
                                                  DELETE FROM [ask_datadb_test].[dbo].[DataImport] 
                                                            WHERE (1=1)
                                                            AND FORMAT(CreateDateTime, 'yyyy-MM-dd') = FORMAT(@date, 'yyyy-MM-dd');
                                             ------------------------------------------------------ [dbo].[File_Download_Log] ---------------------------------------
                                                 DELETE FROM [ask_datadb_test].[dbo].[File_Download_Log] 
                                                 WHERE FORMAT([DownloadDateTime], 'yyyy-MM-dd') = FORMAT(@date, 'yyyy-MM-dd')
                                                          AND [DownloadDateTime] >= @datestart;

                                              ------------------------------------------------------ [dbo].[File_Import_Log] ----------------------------------------
                                                DELETE FROM [ask_datadb_test].[dbo].[File_Import_Log]
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
