using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using ask_read_data.Commons;
using ask_read_data.Controllers;
using ask_read_data.Models;
using ask_read_data.Repository;
using ask_tzn_funamiKD.Commons;

namespace ask_read_data.Servive
{
    public class ExportExcelService : IExportExcel
    {
        public (DataTable, DataTable) GetFloor_Flame_Assy(DateTime dateTime, string bubanType)
        {
            dateTime = Convert.ToDateTime( new DateTime(dateTime.Year, dateTime.Month, dateTime.Day, 00, 00, 00).ToString("yyyy-MM-dd HH:mm:ss"));

            var MyDataTable = new DataTable();
            MyDataTable.Columns.Add("パレットNo", typeof(int));
            MyDataTable.Columns.Add("ラインON", typeof(string));
            MyDataTable.Columns.Add("SEQ", typeof(int));
            MyDataTable.Columns.Add("部品番号", typeof(string));
            MyDataTable.Columns.Add("部品略式記号", typeof(string));
            MyDataTable.Columns.Add("発送予定日", typeof(string));

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
                                                            CommandText = "SP_DataImport_Flame_AssyAndFloor_Assy",
                                                            Connection = connection,
                                                            CommandType = CommandType.StoredProcedure
                                                        };
                    cmd.Parameters.Clear();
                    SqlParameter DateTime = new SqlParameter
                                                            {
                                                                ParameterName = "@DateTime",
                                                                SqlDbType = SqlDbType.DateTime,
                                                                Value = dateTime,
                                                                Direction = ParameterDirection.Input
                                                            };
                    SqlParameter BubanType = new SqlParameter
                                                            {
                                                                ParameterName = "@BubanType",
                                                                SqlDbType = SqlDbType.NVarChar,
                                                                Value = bubanType,
                                                                Direction = ParameterDirection.Input
                                                            };

                    cmd.Parameters.Add(DateTime);
                    cmd.Parameters.Add(BubanType);
                    //SQL実行
                    reader = cmd.ExecuteReader();
                    int PaletNo = 1;
                    int index = 0; 
                    while (reader.Read())
                    {
                        var Row = MyDataTable.NewRow();
                        {
                            index++;
                            Row["パレットNo"] = PaletNo;
                            Row["ラインON"] = Convert.ToDateTime(reader["WAYMD"].ToString()).Year.ToString() + Convert.ToDateTime(reader["WAYMD"].ToString()).Month.ToString() + Convert.ToDateTime(reader["WAYMD"].ToString()).Day.ToString();
                            Row["SEQ"] = Util.NullToBlank((object)reader["SEQ"]);
                            Row["部品番号"] = Util.NullToBlank(reader["BUBAN"].ToString());
                            Row["部品略式記号"] = Util.NullToBlank(reader["KIGO"].ToString());
                            Row["発送予定日"] = (Convert.ToDateTime(reader["FYMD"].ToString()).Year.ToString() + Convert.ToDateTime(reader["FYMD"].ToString()).Month.ToString() + Convert.ToDateTime(reader["FYMD"].ToString()).Day.ToString()).ToString();
                        }
                        MyDataTable.Rows.Add(Row);

                        if ((bubanType == ExportExcelController.FLOOR_ASSY) && (index % 8 == 0))
                        {
                            PaletNo++;
                        }
                        else if ((bubanType == ExportExcelController.FLAME_ASSY) && (index % 4 == 0))
                        {
                            PaletNo++;
                        }
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
            //  Reverse Rows
            DataTable reversedDt = MyDataTable.Clone();
            for (var row = MyDataTable.Rows.Count - 1; row >= 0; row--)
                reversedDt.ImportRow(MyDataTable.Rows[row]);

            return (MyDataTable, reversedDt);

        }

        public int RecordDownloadHistory(DataTable dataTable)
        {
            throw new NotImplementedException();
        }
    }
}
