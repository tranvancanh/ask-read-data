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
        private int ParetoRenban = 0;
        private DateTime LastDateTime = new DateTime();
        private bool isCheckDownload = false;

        private const int MAX_RENBAN_FLOOR_ASSY = 30;
        private const int MAX_RENBAN_FRAME_ASSY = 60;

        public (DataTable, DataTable) GetFloor_Flame_Assy(DateTime dateTime, string bubanType)
        {
            dateTime = Convert.ToDateTime(new DateTime(dateTime.Year, dateTime.Month, dateTime.Day, 00, 00, 00).ToString("yyyy-MM-dd HH:mm:ss"));
            if(!CheckDataExists(dateTime, bubanType))
            {
                return (new DataTable(), new DataTable());
            }
            CreateDateTime = dateTime;
            var result = GetZenkaiDowloadInfor(bubanType, dateTime);
            var lastDateTime = result.Item1;
            var lastPosition = result.Item2;
            var lastParetoRenban = result.Item3;
            lastParetoRenban = CheckRenBan(lastParetoRenban, bubanType);
            var position = new List<string>();
            var MyDataTable = new DataTable();
            DataColumn[] cols ={
                                  new DataColumn("パレットNo",typeof(string)),
                                  new DataColumn("ラインON",typeof(string)),
                                  new DataColumn("SEQ",typeof(string)),
                                  new DataColumn("部品番号",typeof(string)),
                                  new DataColumn("部品略式記号",typeof(string)),
                                  new DataColumn(" ",typeof(string))
                              };

            //DataColumn[] cols ={
            //                      new DataColumn("パレットNo",typeof(string)),
            //                      new DataColumn("ラインON",typeof(string)),
            //                      new DataColumn("SEQ",typeof(string)),
            //                      new DataColumn("部品番号",typeof(string)),
            //                      new DataColumn("部品略式記号",typeof(string)),
            //                      new DataColumn("発送予定日",typeof(string))
            //                      //new DataColumn("",typeof(string))
            //                  };
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
                    int PaletNo = lastParetoRenban + 1;
                    int index = 0;
                    // check  Position

                    while (reader.Read())
                    {
                        if (index == 0)
                        {
                            // CreateDateTime = Convert.ToDateTime(reader["CreateDateTimeForMat"].ToString());
                            /**********************中身のヘッダ*************************/
                            var Row1 = MyDataTable.NewRow();
                            {
                                Row1["パレットNo"] = "パレットNo";
                                Row1["ラインON"] = "ラインON";
                                Row1["SEQ"] = "SEQ";
                                Row1["部品番号"] = "部品番号";
                                Row1["部品略式記号"] = "部品略式記号";
                                Row1[" "] = "";
                            }
                            MyDataTable.Rows.Add(Row1);
                            position.Add("Position");
                        }
                        /**********************中身の体*************************/
                        var Row = MyDataTable.NewRow();
                        {
                            index++;
                            Row["パレットNo"] = PaletNo.ToString();
                            Row["ラインON"] = Convert.ToDateTime(reader["WAYMD"].ToString()).Year.ToString() + Convert.ToDateTime(reader["WAYMD"].ToString()).Month.ToString() + Convert.ToDateTime(reader["WAYMD"].ToString()).Day.ToString();
                            Row["SEQ"] = Util.NullToBlank((object)reader["SEQ"]);
                            Row["部品番号"] = Util.NullToBlank(reader["BUBAN"].ToString());
                            if(bubanType == ExportExcelController.FLOOR_ASSY)
                            {
                                if(Row["部品番号"].ToString() == ExportExcelController.BUHIN_FLOOR_74300WL20P)
                                {
                                    Row["部品略式記号"] = Util.NullToBlank(reader["KIGO"].ToString());
                                    Row[" "] = "";
                                }
                                else if (Row["部品番号"].ToString() == ExportExcelController.BUHIN_FLOOR_74300WL30P)
                                {
                                    Row["部品略式記号"] = "";
                                    Row[" "] = Util.NullToBlank(reader["KIGO"].ToString());
                                }
                            }
                            else if(bubanType == ExportExcelController.FLAME_ASSY)
                            {
                                if(Row["部品番号"].ToString() == ExportExcelController.BUHIN_FLAME_743B2W000P)
                                {
                                    Row["部品略式記号"] = Util.NullToBlank(reader["KIGO"].ToString());
                                    Row[" "] = "";
                                }
                                else if (Row["部品番号"].ToString() == ExportExcelController.BUHIN_FLAME_743B2W010P)
                                {
                                    Row["部品略式記号"] = "";
                                    Row[" "] = Util.NullToBlank(reader["KIGO"].ToString());
                                }
                            }
                            this.ParetoRenban = PaletNo;
                        }
                        position.Add(Util.NullToBlank(reader["Position"].ToString()));
                        MyDataTable.Rows.Add(Row);

                        if ((bubanType == ExportExcelController.FLOOR_ASSY) && (index % ExportExcelController.PALETNO_FLOOR_ASSY == 0))
                        {
                            /**********************中身のヘッダ*************************/
                            Row = MyDataTable.NewRow();
                            {
                                Row["パレットNo"] = "パレットNo";
                                Row["ラインON"] = "ラインON";
                                Row["SEQ"] = "SEQ";
                                Row["部品番号"] = "部品番号";
                                Row["部品略式記号"] = "部品略式記号";
                                Row[" "] = "";
                            }
                            MyDataTable.Rows.Add(Row);
                            position.Add("Position");
                            PaletNo++;
                            if(PaletNo > MAX_RENBAN_FLOOR_ASSY)
                            {
                                PaletNo = 1;
                            }
                        }
                        else if ((bubanType == ExportExcelController.FLAME_ASSY) && (index % ExportExcelController.PALETNO_FLAME_ASSY == 0))
                        {
                            /**********************中身のヘッダ*************************/
                            Row = MyDataTable.NewRow();
                            {
                                Row["パレットNo"] = "パレットNo";
                                Row["ラインON"] = "ラインON";
                                Row["SEQ"] = "SEQ";
                                Row["部品番号"] = "部品番号";
                                Row["部品略式記号"] = "部品略式記号";
                                Row[" "] = "";
                            }
                            MyDataTable.Rows.Add(Row);
                            position.Add("Position");
                            PaletNo++;
                            if (PaletNo > MAX_RENBAN_FRAME_ASSY)
                            {
                                PaletNo = 1;
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
            this.isCheckDownload = true;
            switch (bubanType)
            {
                case ExportExcelController.FLOOR_ASSY:
                    {
                        // 件数 < ExportExcelController.PALETNO_FLOOR_ASSY の場合は、
                        if (MyDataTable.Rows.Count < ExportExcelController.PALETNO_FLOOR_ASSY + 1)
                        {
                            this.LastDateTime = lastDateTime;
                            this.Position = lastPosition;
                            this.ParetoRenban = lastParetoRenban;
                            this.isCheckDownload = false;
                            for (int i = MyDataTable.Rows.Count; i < ExportExcelController.PALETNO_FLOOR_ASSY + 1; i++)
                            {
                                MyDataTable.Rows.Add("", "", "", "", "", "");
                            }
                            /////////////////////////////////////////// Data Reverse //////////////////////////////////////////
                            reversedDt = MyDataTable.Clone();
                            reversedDt.ImportRow(MyDataTable.Rows[0]);
                            for (var row = MyDataTable.Rows.Count - 1; row > 0; row--)
                            { reversedDt.Rows.Add("", "", "", "", "", ""); }
                            break;
                        }
                        this.isCheckDownload = true;
                        Balance = MyDataTable.Rows.Count % (ExportExcelController.PALETNO_FLOOR_ASSY + 1);
                        if (Balance > 0)
                        {
                            for (int i = MyDataTable.Rows.Count - Balance; i < MyDataTable.Rows.Count; i++)
                            {
                                if(MyDataTable.Rows[i]["パレットNo"].ToString() != "パレットNo")
                                {
                                    MyDataTable.Rows[i]["パレットNo"] = ExportExcelController.ASHITA_IKO;
                                }
                            }
                        }
                        //  Get Position 
                        try
                        {
                            var value = Convert.ToInt32(position[MyDataTable.Rows.Count - Balance]);
                            Position = value;
                        }
                        catch (FormatException)
                        {
                            var value = Convert.ToInt32(position[MyDataTable.Rows.Count - Balance - 1]);
                            Position = value;
                        }
                        catch (Exception)
                        {
                            throw new Exception("Position値に問題がありました");
                        }
                        //  欠落している行のデータを追加します , 目的: 9倍数を到達する
                        for (int i = 0; i < 9 - Balance; i++)
                        {
                            MyDataTable.Rows.Add("", "", "", "", "", "");
                        }

                        //////////////////////////////////////// Sheet2のデータを作る //////////////////////////
                        ///// 反逆の場合は、明日以降の件を削除します
                        //var table = MyDataTable.AsEnumerable()
                        //              .Where(r => r.Field<string>("パレットNo") != ExportExcelController.ASHITA_IKO)
                        //              .CopyToDataTable();
                        var table = MyDataTable.Copy();
                        //  明日以降を省略する
                        foreach (DataRow dr in table.Rows) // search whole table
                        {
                            if (dr["パレットNo"].ToString() == ExportExcelController.ASHITA_IKO) // if id==2
                            {
                                dr["パレットNo"] = "";
                                dr["ラインON"] = "";
                                dr["SEQ"] = "";
                                dr["部品番号"] = "";
                                dr["部品略式記号"] = "";
                                dr[" "] = "";
                            }
                        }
                        reversedDt = table.Clone();
                        for (int i = 0; i < table.Rows.Count; i += 9)
                        {
                            if (i % (ExportExcelController.PALETNO_FLOOR_ASSY + 1) == 0) 
                            {
                                var row0 = table.Rows[i];
                                reversedDt.Rows.Add(row0.ItemArray);
                                //continue;
                            };
                            var row1 = table.Rows[i + 1];
                            var row2 = table.Rows[i + 2];
                            var row3 = table.Rows[i + 3];
                            var row4 = table.Rows[i + 4];
                            var row5 = table.Rows[i + 5];
                            var row6 = table.Rows[i + 6];
                            var row7 = table.Rows[i + 7];
                            var row8 = table.Rows[i + 8];
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
                        if (MyDataTable.Rows.Count < ExportExcelController.PALETNO_FLAME_ASSY + 1)
                        {
                            this.LastDateTime = lastDateTime;
                            this.Position = lastPosition;
                            this.ParetoRenban = lastParetoRenban;
                            this.isCheckDownload = false;
                            for (int i = MyDataTable.Rows.Count; i < ExportExcelController.PALETNO_FLAME_ASSY + 1; i++)
                            {
                                MyDataTable.Rows.Add("", "", "", "", "", "");
                            }
                            /////////////////////////////////////////// Data Reverse //////////////////////////////////////////
                            reversedDt = MyDataTable.Clone();
                            reversedDt.ImportRow(MyDataTable.Rows[0]);
                            for (var row = MyDataTable.Rows.Count - 1; row > 0; row--)
                            { reversedDt.Rows.Add("", "", "", "", "", ""); }
                            break;
                        }
                        var ashitaiko = 0;
                        this.isCheckDownload = true;
                        Balance = MyDataTable.Rows.Count % (ExportExcelController.PALETNO_FLAME_ASSY + 1);
                        if (Balance > 0)
                        {
                            //Position = Convert.ToInt32(position[MyDataTable.Rows.Count - Balance - 1]);
                            for (int i = MyDataTable.Rows.Count - Balance; i < MyDataTable.Rows.Count; i++)
                            {
                                if (MyDataTable.Rows[i]["パレットNo"].ToString() != "パレットNo")
                                {
                                    MyDataTable.Rows[i]["パレットNo"] = ExportExcelController.ASHITA_IKO;
                                    ashitaiko++;
                                }
                            }
                        }
                        //  Get Position 
                        try
                        {
                            var value = Convert.ToInt32(position[MyDataTable.Rows.Count - Balance]);
                            Position = value;
                        }
                        catch (FormatException)
                        {
                            var value = Convert.ToInt32(position[MyDataTable.Rows.Count - Balance - 1]);
                            Position = value;
                        }
                        catch (Exception)
                        {
                            throw new Exception("Position値に問題がありました");
                        }
                        //  欠落している行のデータを追加します , 目的: 9倍数を到達する
                        var ceiling = Convert.ToInt32(Math.Ceiling((decimal)MyDataTable.Rows.Count / 5)) * 5;
                        for (int i = MyDataTable.Rows.Count; i < ceiling; i++)
                        {
                            if(i%5 == 0 && i != 0)
                            {
                                MyDataTable.Rows.Add("パレットNo", "ラインON", "SEQ", "部品番号", "部品略式記号", "");
                            }
                            MyDataTable.Rows.Add("", "", "", "", "", "");
                        }
                        //////////////////////////////////////// Sheet2のデータを作る //////////////////////////
                        //var table = MyDataTable.AsEnumerable()
                        //                .Where(r => r.Field<string>("パレットNo") != ExportExcelController.ASHITA_IKO)
                        //                .CopyToDataTable();
                        var table = MyDataTable.Copy();
                        //  明日以降を省略する
                        foreach (DataRow dr in table.Rows) // search whole table
                        {
                            if (dr["パレットNo"].ToString() == ExportExcelController.ASHITA_IKO) // if id==2
                            {
                                dr["パレットNo"] = "";
                                dr["ラインON"] = "";
                                dr["SEQ"] = "";
                                dr["部品番号"] = "";
                                dr["部品略式記号"] = "";
                                dr[" "] = "";
                            }
                        }
                        reversedDt = table.Clone();
                        for (int i = 0; i < table.Rows.Count; i += 5)
                        {
                            if (i % (ExportExcelController.PALETNO_FLAME_ASSY + 1) == 0)
                            {
                                var row0 = table.Rows[i];
                                reversedDt.Rows.Add(row0.ItemArray);
                            };
                            var row1 = table.Rows[i + 1];
                            var row2 = table.Rows[i + 2];
                            var row3 = table.Rows[i + 3];
                            var row4 = table.Rows[i + 4];
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
            //   DataTable reversedDt = MyDataTable.Clone();
            //   for (var row = MyDataTable.Rows.Count - 1; row >= 0; row--)
            //    reversedDt.ImportRow(MyDataTable.Rows[row]);

            return (MyDataTable, reversedDt);

        }
        private int CheckRenBan(int lastParetoRenban, string bubanType)
        {
            switch (bubanType)
            {
                case ExportExcelController.FLOOR_ASSY:
                    {
                        if (lastParetoRenban >= MAX_RENBAN_FLOOR_ASSY)
                        {
                            lastParetoRenban = 0;
                        }
                        break;
                    }
                case ExportExcelController.FLAME_ASSY:
                    {
                        if (lastParetoRenban >= MAX_RENBAN_FRAME_ASSY)
                        {
                            lastParetoRenban = 0;
                        }
                        break;
                    }
                default:
                    {
                        return -9999;
                    }
            }

            return lastParetoRenban;
        }
        private bool CheckDataExists(DateTime dateTime, string bubanType)
        {
            bool isExists = false;
            int statusCode = -500;
            var ConnectionString = new GetConnectString().ConnectionString;
            using (var connection = new SqlConnection(ConnectionString))
            {
                try
                {
                    //Create the command object
                    SqlCommand cmd = new SqlCommand()
                                                    {
                                                        CommandText = "SP_DataImport_CheckData",
                                                        Connection = connection,
                                                        CommandType = CommandType.StoredProcedure
                                                    };
                    //パラメータ初期化
                    cmd.Parameters.Clear();
                    //Set SqlParameter
                    ///////////////////  SetParameter CreateDateTime  ////////////////////////////////////////////
                    SqlParameter CreateDateTime = new SqlParameter
                                                                {
                                                                    ParameterName = "@CreateDateTime",
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

                    SqlParameter StausCode = new SqlParameter
                                                            {
                                                                ParameterName = "@StausCode",
                                                                SqlDbType = SqlDbType.Int,
                                                                Direction = ParameterDirection.Output
                                                            };
                    cmd.Parameters.Add(CreateDateTime);
                    cmd.Parameters.Add(BubanType);
                    cmd.Parameters.Add(StausCode);

                    connection.Open();
                    cmd.ExecuteNonQuery();
                    statusCode = Convert.ToInt32(StausCode.Value);
                }
                catch
                {
                    throw;
                }
                finally
                {
                    //close connection
                    connection.Close();
                    // connection dispose
                    connection.Dispose();
                }
            }

            if (statusCode == 200)
            {
                isExists = true;
            }
            else
            {
                isExists = false;
            }

            return isExists;
        }
        public int RecordDownloadHistory(ref DataTable dataTable, string bubanType, List<Claim> Claims)
        {
            var lastDownloadDateTime = new DateTime();
            if (isCheckDownload)
            {
                lastDownloadDateTime = CreateDateTime;
            }
            else
            {
                lastDownloadDateTime = this.LastDateTime;
            }
            var UserName = Claims.Where(c => c.Type == ClaimTypes.Name).First().Value;
            int Position = this.Position;
            var paretoRenban = 0;
            int statusCode = 0;
            int affectedRows = 0;
            var ConnectionString = new GetConnectString().ConnectionString;
            using (var connection = new SqlConnection(ConnectionString))
            {
                try
                {
                    bool isCheckRenban = false;
                    // パレット連番の取得
                    for(int i = dataTable.Rows.Count -1; i >= 0; i--)
                    {
                        var value = dataTable.Rows[i]["パレットNo"].ToString();
                        if(Util.NullToBlank(value) != string.Empty && value != "パレットNo" && value != ExportExcelController.ASHITA_IKO)
                        {
                            paretoRenban = Convert.ToInt32(value);
                            isCheckRenban = true;
                            break;
                        }
                    }
                    if(isCheckRenban != true) { paretoRenban = this.ParetoRenban; }
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
                    SqlParameter ParetoRenban = new SqlParameter
                                                            {
                                                                ParameterName = "@ParetoRenban",
                                                                SqlDbType = SqlDbType.Int,
                                                                Value = paretoRenban,
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
                    cmd.Parameters.Add(ParetoRenban);
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

        private (DateTime, int, int) GetZenkaiDowloadInfor(string bubanMeiType, DateTime dateTime)
        {
            var LastDownloadDateTime = new DateTime(1900, 01, 01, 00, 00, 00);
            int lastPosition = 0;
            int lastParetoRenban = 0;
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
                        lastParetoRenban  = Util.NullToBlank((object)reader["ParetoRenban"]);
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

            return (LastDownloadDateTime, lastPosition, lastParetoRenban);
        }
    }
}
