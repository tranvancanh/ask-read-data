﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using ask_read_data.Commons;
using ask_read_data.Controllers;
using ask_read_data.Dao;
using ask_read_data.Models;
using ask_read_data.Models.Entity;
using ask_read_data.Models.ViewModel;
using ask_read_data.Repository;
using ask_tzn_funamiKD.Commons;

namespace ask_read_data.Servive
{
    public class ExportExcelService : IExportExcel
    {
        private DateTime DateTime = new DateTime();
        private DateTime LastDownloadDateTime = new DateTime();
        private int Position = 0;
        private int ParetoRenban = 0;
        private bool isDownloadStatus = false;

        private const int MAX_RENBAN_FLOOR_ASSY = 30;
        private const int MAX_RENBAN_FRAME_ASSY = 60;

        public (DataTable, DataTable) GetFloor_Flame_Assy(ExportExcelViewModel modelRequset, string bubanType)
        {
            var dateStart = new DateTime();
            var startPosition = 0;
            var startParetoRenban = 0;
            switch (bubanType)
            {
                case ExportExcelController.FL00R_ASSY:
                    {
                        dateStart = Convert.ToDateTime(new DateTime(modelRequset.Floor_Assy.Year, modelRequset.Floor_Assy.Month, modelRequset.Floor_Assy.Day, 00, 00, 00).ToString("yyyy-MM-dd HH:mm:ss"));
                        startPosition = modelRequset.Floor_Position;
                        startParetoRenban = modelRequset.Floor_ParetoRenban;
                        if (startParetoRenban >= MAX_RENBAN_FLOOR_ASSY || startParetoRenban <= 0)
                            startParetoRenban = 0;
                        break;
                    }
                case ExportExcelController.FRAME_ASSY:
                    {
                        dateStart = Convert.ToDateTime(new DateTime(modelRequset.Flame_Assy.Year, modelRequset.Flame_Assy.Month, modelRequset.Flame_Assy.Day, 00, 00, 00).ToString("yyyy-MM-dd HH:mm:ss"));
                        startPosition = modelRequset.Flame_Position;
                        startParetoRenban = modelRequset.Flame_ParetoRenban;
                        if (startParetoRenban >= MAX_RENBAN_FRAME_ASSY || startParetoRenban <= 0)
                            startParetoRenban = 0;
                        break;
                    }
            }
            this.DateTime = dateStart;
            //CreateDateTime = dateTime;
            //var result = GetZenkaiDowloadInfor(bubanType, dateTime);
            //var lastDateTime = result.Item1;
            //var lastPosition = result.Item2;
            //var lastParetoRenban = result.Item3;
            //lastParetoRenban = CheckRenBan(lastParetoRenban, bubanType);
            //var dateStart = dateTime;
            //var position = new List<string>();
            var temporaryData = new List<TemporaryData>();
            var MyDataTable = new DataTable();
            DataColumn[] cols ={
                                  new DataColumn("パレットNo",typeof(string)),
                                  new DataColumn("ラインON",typeof(string)),
                                  new DataColumn("SEQ",typeof(string)),
                                  new DataColumn("部品番号",typeof(string)),
                                  new DataColumn("部品略式記号",typeof(string)),
                                  new DataColumn(" ",typeof(string))
                              };
            MyDataTable.Columns.AddRange(cols);

            var ConnectionString = new GetConnectString().ConnectionString();
            using (var connection = new SqlConnection(ConnectionString))
            {
                SqlDataReader reader = null;
                try
                {
                    connection.Open();
                    //Create the command object
                    SqlCommand cmd = new SqlCommand()
                                                                {
                                                                    CommandText = ExportExcelDao.SP_DataImport_Flame_AssyAndFloor_Assy(),
                                                                    Connection = connection,
                                                                    CommandType = CommandType.Text
                                                                };
                    cmd.Parameters.Clear();
                    SqlParameter BubanType = new SqlParameter
                                                                {
                                                                    ParameterName = "@BubanType",
                                                                    SqlDbType = SqlDbType.NVarChar,
                                                                    Value = bubanType,
                                                                    Direction = ParameterDirection.Input
                                                                };

                    //////////////////////////////////////////////////////////  前回残りをチェック  /////////////////////////////////////////////////////////////
                    SqlParameter StartDateTime = new SqlParameter
                                                                {
                                                                    ParameterName = "@StartDateTime",
                                                                    SqlDbType = SqlDbType.DateTime,
                                                                    Value = dateStart,
                                                                    Direction = ParameterDirection.Input
                                                                };

                    SqlParameter StartPosition = new SqlParameter
                                                                {
                                                                    ParameterName = "@StartPosition",
                                                                    SqlDbType = SqlDbType.Int,
                                                                    Value = startPosition,
                                                                    Direction = ParameterDirection.Input
                                                                };

                    cmd.Parameters.Add(BubanType);
                    cmd.Parameters.Add(StartDateTime);
                    cmd.Parameters.Add(StartPosition);
                    //SQL実行
                    reader = cmd.ExecuteReader();
                    int PaletNo = startParetoRenban + 1;
                    int index = 0;
                    if (!reader.HasRows)
                    {
                        return (new DataTable(), new DataTable());
                    }
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
                            temporaryData.Add(new TemporaryData());
                        }
                        /**********************中身の体*************************/
                        var Row = MyDataTable.NewRow();
                        {
                            index++;
                            Row["パレットNo"] = PaletNo.ToString();
                            Row["ラインON"] = Convert.ToDateTime(reader["WAYMD"].ToString()).Year.ToString() + Convert.ToDateTime(reader["WAYMD"].ToString()).Month.ToString() + Convert.ToDateTime(reader["WAYMD"].ToString()).Day.ToString();
                            Row["SEQ"] = Util.NullToBlank((object)reader["SEQ"]);
                            Row["部品番号"] = Util.NullToBlank(reader["BUBAN"].ToString());
                            if(bubanType == ExportExcelController.FL00R_ASSY)
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
                            else if(bubanType == ExportExcelController.FRAME_ASSY)
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
                        temporaryData.Add(new TemporaryData() { Position = Util.NullToBlank(reader["Position"].ToString()), CreateDateTime = Convert.ToDateTime(reader["CreateDateTime"].ToString()) });
                        MyDataTable.Rows.Add(Row);

                        if ((bubanType == ExportExcelController.FL00R_ASSY) && (index % ExportExcelController.PALETNO_FLOOR_ASSY == 0))
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
                            temporaryData.Add(new TemporaryData());
                            PaletNo++;
                            if(PaletNo > MAX_RENBAN_FLOOR_ASSY)
                            {
                                PaletNo = 1;
                            }
                        }
                        else if ((bubanType == ExportExcelController.FRAME_ASSY) && (index % ExportExcelController.PALETNO_FLAME_ASSY == 0))
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
                            temporaryData.Add(new TemporaryData());
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
            this.isDownloadStatus = true;
            switch (bubanType)
            {
                case ExportExcelController.FL00R_ASSY:
                    {
                        // 件数 < ExportExcelController.PALETNO_FLOOR_ASSY の場合は、
                        if (MyDataTable.Rows.Count < ExportExcelController.PALETNO_FLOOR_ASSY + 1)
                        {
                            this.Position = startPosition;
                            this.ParetoRenban = startParetoRenban;
                            this.isDownloadStatus = false;
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
                        this.isDownloadStatus = true;
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
                        var index = MyDataTable.Rows.Count - Balance;
                        bool isConvert = false;
                        for(int i = index; i >= 0; i--)
                        {
                            var value = temporaryData[i];
                            if(int.TryParse(value.Position.ToString(), out int j))
                            {
                                this.Position = j;
                                this.LastDownloadDateTime = Convert.ToDateTime(value.CreateDateTime.ToString("yyyy-MM-dd"));
                                isConvert = true;
                                break;
                            }
                        }
                        if(isConvert != true) 
                        { 
                            this.Position = -1;
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
                case ExportExcelController.FRAME_ASSY:
                    {
                        // 件数 < ExportExcelController.PALETNO_FLAME_ASSY の場合は、
                        if (MyDataTable.Rows.Count < ExportExcelController.PALETNO_FLAME_ASSY + 1)
                        {
                            this.Position = startPosition;
                            this.ParetoRenban = startParetoRenban;
                            this.isDownloadStatus = false;
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
                        this.isDownloadStatus = true;
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
                        var index = MyDataTable.Rows.Count - Balance;
                        bool isConvert = false;
                        for (int i = index; i >= 0; i--)
                        {
                            var value = temporaryData[i];
                            if (int.TryParse(value.Position.ToString(), out int j))
                            {
                                this.Position = j;
                                this.LastDownloadDateTime = Convert.ToDateTime(value.CreateDateTime.ToString("yyyy-MM-dd"));
                                isConvert = true;
                                break;
                            }
                        }
                        if (isConvert != true)
                        {
                            this.Position = -1;
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
                case ExportExcelController.FL00R_ASSY:
                    {
                        if (lastParetoRenban >= MAX_RENBAN_FLOOR_ASSY)
                        {
                            lastParetoRenban = 0;
                        }
                        break;
                    }
                case ExportExcelController.FRAME_ASSY:
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
                try
                {
                //Create the command object
                isExists = ExportExcelDao.SP_DataImport_CheckData(dateTime, bubanType);
                }
                catch
                {
                    throw;
                }
 
            return isExists;
        }
        public int RecordDownloadHistory(ref DataTable dataTable, string bubanType, List<Claim> Claims)
        {
            var UserName = Claims.Where(c => c.Type == ClaimTypes.Name).First().Value;
            var lastDownloadDateTime = new DateTime();
            var position = 0;
            var paretoRenban = 0;
            if (this.isDownloadStatus)
            {
                lastDownloadDateTime = this.LastDownloadDateTime;
                position = this.Position;
                bool isCheckRenban = false;
                // パレット連番の取得
                for (int i = dataTable.Rows.Count - 1; i >= 0; i--)
                {
                    var value = dataTable.Rows[i]["パレットNo"].ToString();
                    if (Util.NullToBlank(value) != string.Empty && value != "パレットNo" && value != ExportExcelController.ASHITA_IKO)
                    {
                        paretoRenban = Convert.ToInt32(value);
                        isCheckRenban = true;
                        break;
                    }
                }
                if (isCheckRenban != true) { throw new Exception("パレット連番がおかしいです"); }
            }
            else
            {
                var result = GetZenkaiDowloadInfor(bubanType, DateTime.Today);
                lastDownloadDateTime = result.Item1;
                position = result.Item2;
                paretoRenban = result.Item3;
            }
            int statusCode = 0;
            int affectedRows = 0;
            var ConnectionString = new GetConnectString().ConnectionString();
            using (var connection = new SqlConnection(ConnectionString))
            {
                try
                {
                    
                    //Create the command object
                    SqlCommand cmd = new SqlCommand()
                                                        {
                                                            CommandText = ExportExcelDao.SP_File_Download_Log_Insert(),
                                                            Connection = connection,
                                                            CommandType = CommandType.Text
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
                    SqlParameter Position = new SqlParameter
                                                            {
                                                                ParameterName = "@Position",
                                                                SqlDbType = SqlDbType.Int,
                                                                Value = position,
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
                    cmd.Parameters.Add(Position);
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
            var ConnectionString = new GetConnectString().ConnectionString();
            using (var connection = new SqlConnection(ConnectionString))
            {
                SqlDataReader reader = null;
                try
                {
                    //Create the command object
                    SqlCommand cmd = new SqlCommand()
                                                    {
                                                        CommandText = ExportExcelDao.SP_GetLastDowloadInfo(),
                                                        Connection = connection,
                                                        CommandType = CommandType.Text
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
                        LastDownloadDateTime = Convert.ToDateTime(Convert.ToDateTime(reader["LastDownloadDateTime"].ToString()).ToString("yyyy-MM-dd"));
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

        public Tuple<int, int> FindPositionParetoRenban(DateTime date, string bubanType)
        {
            try
            {
                var result = ExportExcelDao.GetPositonParetoRenban_Before_Download(date, bubanType);
                var position = result.Item1;
                var renban = result.Item2;

                return result;
            }
            catch
            {
                throw;
            }
           
        }

        public List<DataModel> FindRemainingDataOfLastTime(ExportExcelViewModel viewModel)
        {
            try
            {
                return new List<DataModel>();
                //return ExportExcelDao.GetRemainingDataOfLastTime(viewModel);
            }
            catch
            {
                throw;
            }
        }

        public List<FileDownloadLogModel> FindDownloadHistory(DateTime date)
        {
            try
            {
                return ExportExcelDao.GetDownloadHistory(date);
            }
            catch
            {
                throw;
            }
        }

        public Tuple<DateTime, int, int> FindPositionParetoRenbanLasttime(string bubanType)
        {
            try
            {
                return ExportExcelDao.GetPositionParetoRenbanLasttime(bubanType);
            }
            catch
            {
                throw;
            }
        }
    }
}
