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
            dateTime = Convert.ToDateTime( new DateTime(dateTime.Year, dateTime.Month, dateTime.Day, 00, 00, 00).ToString("yyyy-MM-dd HH:mm:ss"));
            int lastPosition = GetPosition(bubanType, dateTime);
            var position = new List<int>();
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

                    SqlParameter Position = new SqlParameter
                                                            {
                                                                ParameterName = "@Position",
                                                                SqlDbType = SqlDbType.Int,
                                                                Value = lastPosition,
                                                                Direction = ParameterDirection.Input
                                                            };

                    cmd.Parameters.Add(DateTime);
                    cmd.Parameters.Add(BubanType);
                    cmd.Parameters.Add(Position);
                    //SQL実行
                    reader = cmd.ExecuteReader();
                    int PaletNo = 1;
                    int index = 0;
                    // check  Position
                    
                    while (reader.Read())
                    {
                        if(index == 0) CreateDateTime = Convert.ToDateTime(reader["CreateDateTimeForMat"].ToString());
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
            // Positionの取得
            int Balance = 0;
            switch (bubanType)
            {
                case ExportExcelController.FLOOR_ASSY:
                    {
                        if(MyDataTable.Rows.Count < ExportExcelController.PALETNO_FLOOR_ASSY)
                        {
                            Position = 0;
                            break;
                        }
                        Balance = MyDataTable.Rows.Count % ExportExcelController.PALETNO_FLOOR_ASSY;
                        if (Balance > 0)
                            Position = Convert.ToInt32(position[MyDataTable.Rows.Count - Balance -1]);
                        break;
                    }
                case ExportExcelController.FLAME_ASSY:
                    {

                        if (MyDataTable.Rows.Count < ExportExcelController.PALETNO_FLAME_ASSY)
                        {
                            Position = 0;
                            break;
                        }
                        Balance = MyDataTable.Rows.Count % ExportExcelController.PALETNO_FLAME_ASSY;
                        if (Balance > 0)
                            Position = Convert.ToInt32(position[MyDataTable.Rows.Count - Balance - 1]);
                        break;
                    }
                default:
                    {
                        Balance = -1;
                        throw new Exception("部品番号名称がおかしいです");
                    }
            }
            //  Reverse Rows
            DataTable reversedDt = MyDataTable.Clone();
            for (var row = MyDataTable.Rows.Count - 1; row >= 0; row--)
                reversedDt.ImportRow(MyDataTable.Rows[row]);

            return (MyDataTable, reversedDt);

        }

        public int RecordDownloadHistory(ref DataTable dataTable, string bubanType, List<Claim> Claims)
        {
            DateTime LastImportDateTime = CreateDateTime;
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
                    SqlParameter LastImportDateTime1 = new SqlParameter
                                                                {
                                                                    ParameterName = "@LastImportDateTime",
                                                                    SqlDbType = SqlDbType.DateTime,
                                                                    Value = LastImportDateTime,
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
                    cmd.Parameters.Add(LastImportDateTime1);
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
                    if(connection != null)
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

        private int GetPosition(string bubanMeiType, DateTime dateTime)
        {
            int lastPosition = 0;
            var ConnectionString = new GetConnectString().ConnectionString;
            using (var connection = new SqlConnection(ConnectionString))
            {
                try
                {
                    //Create the command object
                    SqlCommand cmd = new SqlCommand()
                                                        {
                                                            CommandText = "SP_GetPosition",
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

                    ///////////////////  SetParameter StausCode  ////////////////////////////////////////////
                    SqlParameter StausCode = new SqlParameter
                                                        {
                                                            ParameterName = "@StausCode",
                                                            SqlDbType = SqlDbType.Int,
                                                            Direction = ParameterDirection.Output
                                                        };
                    cmd.Parameters.Add(BubanMeiType);
                    cmd.Parameters.Add(DateTime);
                    cmd.Parameters.Add(StausCode);

                    var value = cmd.ExecuteScalar();
                    lastPosition = Convert.ToInt32(value);
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


            return lastPosition;
        }
    }
}
