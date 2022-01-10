using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using ask_read_data.Commons;

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
                if(connection != null) { connection.Close(); }
                if(command != null) { connection.Dispose(); }
            }
            return position;
        }
    }
}
