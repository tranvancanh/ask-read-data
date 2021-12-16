using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Security.Claims;
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
        private DateTime CreateDateTime;
        private int Position = 0;
        public (DataTable, DataTable) GetFloor_Flame_Assy(DateTime dateTime, string bubanType)
        {
            dateTime = Convert.ToDateTime(new DateTime(dateTime.Year, dateTime.Month, dateTime.Day, 00, 00, 00).ToString("yyyy-MM-dd HH:mm:ss"));
            var result = GetPosition(bubanType, dateTime);
            var lastDateTime = result.Item1;
            var lastPosition = result.Item2;
            var position = new List<int>();
            var MyDataTable = new DataTable();
            DataColumn[] cols ={
                                  new DataColumn("No",typeof(string)),
                                  new DataColumn("MMC生産日",typeof(string)),
                                  new DataColumn("SEQ",typeof(string)),
                                  new DataColumn("パレットNo",typeof(string)),
                                  new DataColumn("部品番号",typeof(string)),
                                  new DataColumn("部品略式記号",typeof(string)),
                                  new DataColumn("",typeof(string))
                              };
            MyDataTable.Columns.AddRange(cols);

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

                    //////////////////////////////////////////////////////////  前回残りをチェック  /////////////////////////////////////////////////////////////
                    SqlParameter LastDateTime = new SqlParameter
                    {
                        ParameterName = "@LastDateTime",
                        SqlDbType = SqlDbType.DateTime,
                        Value = lastDateTime,
                        Direction = ParameterDirection.Input
                    };

                    SqlParameter LastPosition = new SqlParameter
                    {
                        ParameterName = "@LastPosition",
                        SqlDbType = SqlDbType.Int,
                        Value = lastPosition,
                        Direction = ParameterDirection.Input
                    };

                    cmd.Parameters.Add(DateTime);
                    cmd.Parameters.Add(BubanType);
                    cmd.Parameters.Add(LastDateTime);
                    cmd.Parameters.Add(LastPosition);
                    //SQL実行
                    reader = cmd.ExecuteReader();
                    int PaletNo = 1;
                    int index = 0;
                    // check  Position

                    while (reader.Read())
                    {
                        if (index == 0) CreateDateTime = Convert.ToDateTime(reader["CreateDateTimeForMat"].ToString());
                        var Row = MyDataTable.NewRow();
                        {
                            index++;
                            Row["パレットNo"] = PaletNo.ToString();
                            Row["ラインON"] = Convert.ToDateTime(reader["WAYMD"].ToString()).Year.ToString() + Convert.ToDateTime(reader["WAYMD"].ToString()).Month.ToString() + Convert.ToDateTime(reader["WAYMD"].ToString()).Day.ToString();
                            Row["SEQ"] = Util.NullToBlank((object)reader["SEQ"]);
                            Row["部品番号"] = Util.NullToBlank(reader["BUBAN"].ToString());
                            Row["部品略式記号"] = Util.NullToBlank(reader["KIGO"].ToString());
                            Row["発送予定日"] = (Convert.ToDateTime(reader["FYMD"].ToString()).Year.ToString() + Convert.ToDateTime(reader["FYMD"].ToString()).Month.ToString() + Convert.ToDateTime(reader["FYMD"].ToString()).Day.ToString()).ToString();
                        }
                        position.Add(Util.NullToBlank((object)reader["Position"]));
                        MyDataTable.Rows.Add(Row);

                        if ((bubanType == ExportExcelController.FLOOR_ASSY) && (index % ExportExcelController.PALETNO_FLOOR_ASSY == 0))
                        {
                            PaletNo++;
                        }
                        else if ((bubanType == ExportExcelController.FLAME_ASSY) && (index % ExportExcelController.PALETNO_FLAME_ASSY == 0))
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
            DataTable reversedDt = null;
            // Sheet1 と Sheet2は分けること
            // Positionの取得
            int Balance = 0;
            switch (bubanType)
            {
                case ExportExcelController.FLOOR_ASSY:
                    {
                        // 件数 < ExportExcelController.PALETNO_FLOOR_ASSY の場合は、
                        if (MyDataTable.Rows.Count < ExportExcelController.PALETNO_FLOOR_ASSY)
                        {
                            Position = 0;
                            reversedDt = MyDataTable.Clone();
                            for (var row = MyDataTable.Rows.Count - 1; row >= 0; row--)
                                reversedDt.ImportRow(MyDataTable.Rows[row]);
                            break;
                        }
                        Balance = MyDataTable.Rows.Count % ExportExcelController.PALETNO_FLOOR_ASSY;
                        if (Balance > 0)
                        {
                            Position = Convert.ToInt32(position[MyDataTable.Rows.Count - Balance - 1]);
                            for (int i = MyDataTable.Rows.Count - Balance; i < MyDataTable.Rows.Count; i++)
                            {
                                MyDataTable.Rows[i]["パレットNo"] = ExportExcelController.ASHITA_IKO;
                            }
                        }
                        //////////////////////////////////////// Sheet2のデータ //////////////////////////
                        ///// 反逆の場合は、明日以降の件を削除します
                        var table = MyDataTable.AsEnumerable()
                                      .Where(r => r.Field<string>("パレットNo") != ExportExcelController.ASHITA_IKO)
                                      .CopyToDataTable();
                        reversedDt = table.Clone();
                        for (int i = 0; i < table.Rows.Count; i += 8)
                        {
                            var row1 = table.Rows[i + 0];
                            var row2 = table.Rows[i + 1];
                            var row3 = table.Rows[i + 2];
                            var row4 = table.Rows[i + 3];
                            var row5 = table.Rows[i + 4];
                            var row6 = table.Rows[i + 5];
                            var row7 = table.Rows[i + 6];
                            var row8 = table.Rows[i + 7];
                            reversedDt.Rows.Add(row8.ItemArray);
                            reversedDt.Rows.Add(row7.ItemArray);
                            reversedDt.Rows.Add(row6.ItemArray);
                            reversedDt.Rows.Add(row5.ItemArray);
                            reversedDt.Rows.Add(row4.ItemArray);
                            reversedDt.Rows.Add(row3.ItemArray);
                            reversedDt.Rows.Add(row2.ItemArray);
                            reversedDt.Rows.Add(row1.ItemArray);
                        }
                        break;
                    }
                case ExportExcelController.FLAME_ASSY:
                    {
                        // 件数 < ExportExcelController.PALETNO_FLAME_ASSY の場合は、
                        if (MyDataTable.Rows.Count < ExportExcelController.PALETNO_FLAME_ASSY)
                        {
                            Position = 0;
                            reversedDt = MyDataTable.Clone();
                            for (var row = MyDataTable.Rows.Count - 1; row >= 0; row--)
                                reversedDt.ImportRow(MyDataTable.Rows[row]);
                            break;
                        }
                        Balance = MyDataTable.Rows.Count % ExportExcelController.PALETNO_FLAME_ASSY;
                        if (Balance > 0)
                        {
                            Position = Convert.ToInt32(position[MyDataTable.Rows.Count - Balance - 1]);
                            for (int i = MyDataTable.Rows.Count - Balance; i < MyDataTable.Rows.Count; i++)
                            {
                                MyDataTable.Rows[i]["パレットNo"] = ExportExcelController.ASHITA_IKO;
                            }
                        }
                        //////////////////////////////////////// Sheet2のデータ //////////////////////////
                        var table = MyDataTable.AsEnumerable()
                                        .Where(r => r.Field<string>("パレットNo") != ExportExcelController.ASHITA_IKO)
                                        .CopyToDataTable();
                        reversedDt = table.Clone();
                        for (int i = 0; i < table.Rows.Count; i += 4)
                        {
                            var row1 = table.Rows[i + 0];
                            var row2 = table.Rows[i + 1];
                            var row3 = table.Rows[i + 2];
                            var row4 = table.Rows[i + 3];
                            reversedDt.Rows.Add(row4.ItemArray);
                            reversedDt.Rows.Add(row3.ItemArray);
                            reversedDt.Rows.Add(row2.ItemArray);
                            reversedDt.Rows.Add(row1.ItemArray);
                        }
                        break;
                    }
                default:
                    {
                        Balance = -1;
                        throw new Exception("部品番号名称がおかしいです");
                    }
            }

            ////  Reverse Rows
            //DataTable reversedDt = MyDataTable.Clone();
            //for (var row = MyDataTable.Rows.Count - 1; row >= 0; row--)
            //    reversedDt.ImportRow(MyDataTable.Rows[row]);

            return (MyDataTable, reversedDt);

        }

        public int RecordDownloadHistory(ref DataTable dataTable, string bubanType, List<Claim> Claims)
        {
            DateTime lastDownloadDateTime = CreateDateTime;
            var UserName = Claims.Where(c => c.Type == ClaimTypes.Name).First().Value;
            int Position = this.Position;
            int statusCode = 0;
            int affectedRows = 0;
            var ConnectionString = new GetConnectString().ConnectionString;
            using (var connection = new SqlConnection(ConnectionString))
            {
                try
                {
                    //Create the command object
                    SqlCommand cmd = new SqlCommand()
                    {
                        CommandText = "SP_File_Download_Log_Insert",
                        Connection = connection,
                        CommandType = CommandType.StoredProcedure
                    };
                    //パラメータ初期化
                    cmd.Parameters.Clear();
                    connection.Open();
                    //Set SqlParameter
                    ///////////////////  SetParameter UserName  ////////////////////////////////////////////
                    SqlParameter BubanMeiType = new SqlParameter
                    {
                        ParameterName = "@BubanMeiType",
                        SqlDbType = SqlDbType.NVarChar,
                        Value = bubanType,
                        Direction = ParameterDirection.Input
                    };
                    ///////////////////  SetParameter Password  ////////////////////////////////////////////
                    SqlParameter LastDownloadDateTime = new SqlParameter
                    {
                        ParameterName = "@LastDownloadDateTime",
                        SqlDbType = SqlDbType.DateTime,
                        Value = lastDownloadDateTime,
                        Direction = ParameterDirection.Input
                    };
                    SqlParameter Position1 = new SqlParameter
                    {
                        ParameterName = "@Position",
                        SqlDbType = SqlDbType.Int,
                        Value = Position,
                        Direction = ParameterDirection.Input
                    };

                    SqlParameter User = new SqlParameter
                    {
                        ParameterName = "@User",
                        SqlDbType = SqlDbType.NVarChar,
                        Value = UserName,
                        Direction = ParameterDirection.Input
                    };

                    ///////////////////  SetParameter StausCode  ////////////////////////////////////////////
                    SqlParameter StausCode = new SqlParameter
                    {
                        ParameterName = "@StausCode",
                        SqlDbType = SqlDbType.Int,
                        Direction = ParameterDirection.Output
                    };
                    cmd.Parameters.Add(BubanMeiType);
                    cmd.Parameters.Add(LastDownloadDateTime);
                    cmd.Parameters.Add(Position1);
                    cmd.Parameters.Add(User);
                    cmd.Parameters.Add(StausCode);

                    affectedRows = cmd.ExecuteNonQuery();
                    statusCode = Convert.ToInt32(StausCode.Value);
                }
                catch
                {
                    throw;
                }
                finally
                {
                    if (connection != null)
                    {
                        //close connection
                        connection.Close();
                        // connection dispose
                        connection.Dispose();
                    }

                }
            }

            if (affectedRows > 0 && statusCode == 200)
                return affectedRows;
            else return -1;
        }

        private (DateTime, int) GetPosition(string bubanMeiType, DateTime dateTime)
        {
            var LastDownloadDateTime = new DateTime(1900, 01, 01, 00, 00, 00);
            int lastPosition = 0;
            var ConnectionString = new GetConnectString().ConnectionString;
            using (var connection = new SqlConnection(ConnectionString))
            {
                SqlDataReader reader = null;
                try
                {
                    //Create the command object
                    SqlCommand cmd = new SqlCommand()
                    {
                        CommandText = "SP_GetLastDowloadInfo",
                        Connection = connection,
                        CommandType = CommandType.StoredProcedure
                    };
                    //パラメータ初期化
                    cmd.Parameters.Clear();
                    connection.Open();
                    //Set SqlParameter
                    ///////////////////  SetParameter UserName  ////////////////////////////////////////////
                    SqlParameter BubanMeiType = new SqlParameter
                    {
                        ParameterName = "@BubanMeiType",
                        SqlDbType = SqlDbType.NVarChar,
                        Value = bubanMeiType,
                        Direction = ParameterDirection.Input
                    };
                    ///////////////////  SetParameter Password  ////////////////////////////////////////////
                    SqlParameter DateTime = new SqlParameter
                    {
                        ParameterName = "@DateTime",
                        SqlDbType = SqlDbType.DateTime,
                        Value = dateTime,
                        Direction = ParameterDirection.Input
                    };


                    cmd.Parameters.Add(BubanMeiType);
                    cmd.Parameters.Add(DateTime);

                    reader = cmd.ExecuteReader();
                    while (reader.Read())
                    {
                        LastDownloadDateTime = Convert.ToDateTime(reader["LastDownloadDateTime"].ToString());
                        lastPosition = Util.NullToBlank((object)reader["Position"]);
                        break;
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
                        //close connection
                        connection.Close();
                        // connection dispose
                        connection.Dispose();
                    }
                    if (reader != null)
                    {
                        reader.Close();
                    }

                }
            }

            return (LastDownloadDateTime, lastPosition);
        }
    }
}
