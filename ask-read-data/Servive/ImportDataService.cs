using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ask_read_data.Repository;
using ask_read_data.Models;
using ask_read_data.Commons;
using System.Data.SqlClient;
using System.Data;
using System.Security.Claims;
using ask_read_data.Models.Entity;
using ask_read_data.Dao;
using ask_read_data.Models.ViewModel;

namespace ask_read_data.Servive
{
    public class ImportDataService : IImportData
    {
     
        public ResponResult ImportDataDB(List<object> datas1, List<Claim> Claims, List<Bu_MastarModel> buMastars)
        {
            List<DataModel> datas = new List<DataModel>();
            foreach (var data in datas1)
            {
                DataModel item = (DataModel)data;
                datas.Add(item);
            }
            var listFileName = new List<string>();
            var respon = new ResponResult();
            int affectedRows = 0;
            var ConnectionString = new GetConnectString().ConnectionString;
            var UserName = Claims.Where(c => c.Type == ClaimTypes.Name).First().Value;
            DateTime insertDateTime = Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            // インポートの前に、Positionの値をチェックする
            var position = ImportDataDao.DataImport_Before_CheckPosition(insertDateTime);
            using (var connection = new SqlConnection(ConnectionString))
            {
                int lineNo = 0;
                SqlTransaction transaction = null;
                 
                connection.Open();
                transaction = connection.BeginTransaction();
                try 
                {
                    //Create the command object
                    SqlCommand cmd = new SqlCommand()
                                                        {
                                                            CommandText = ImportDataDao.SP_DataImportInsert(),
                                                            Connection = connection,
                                                            CommandType = CommandType.Text,
                                                            Transaction = transaction
                                                        };
                    foreach (var data in datas)
                    {
                        // Get all file name
                        if(!listFileName.Contains(data.FileName))
                        {
                            listFileName.Add(data.FileName);
                        }
                        lineNo = data.LineNumber;
                        //パラメータ初期化
                        cmd.Parameters.Clear();
                        //Set SqlParameter
                        ///////////////////  SetParameter WAYMD  ////////////////////////////////////////////
                        SqlParameter WAYMD = new SqlParameter
                                                                 {
                                                                     ParameterName = "@WAYMD",
                                                                     SqlDbType = SqlDbType.DateTime,
                                                                     Value = data.WAYMD,
                                                                     Direction = ParameterDirection.Input
                                                                 };

                        ///////////////////  SetParameter SEQ  ////////////////////////////////////////////
                        SqlParameter SEQ = new SqlParameter
                                                                 {
                                                                     ParameterName = "@SEQ",
                                                                     SqlDbType = SqlDbType.Int,
                                                                     Value = data.SEQ,
                                                                     Direction = ParameterDirection.Input
                                                                 };

                        ///////////////////  SetParameter KATASIKI  ////////////////////////////////////////////
                        SqlParameter KATASIKI = new SqlParameter
                                                                 {
                                                                     ParameterName = "@KATASIKI",
                                                                     SqlDbType = SqlDbType.NVarChar,
                                                                     Value = data.KATASIKI,
                                                                     Direction = ParameterDirection.Input
                                                                 };

                        ///////////////////  SetParameter MEISHO  ////////////////////////////////////////////
                        SqlParameter MEISHO = new SqlParameter
                                                                 {
                                                                     ParameterName = "@MEISHO",
                                                                     SqlDbType = SqlDbType.NVarChar,
                                                                     Value = data.MEISHO,
                                                                     Direction = ParameterDirection.Input
                                                                 };

                        ///////////////////  SetParameter FILLER1  ////////////////////////////////////////////
                        SqlParameter FILLER1 = new SqlParameter
                                                                {
                                                                    ParameterName = "@FILLER1",
                                                                    SqlDbType = SqlDbType.NVarChar,
                                                                    Value = data.FILLER1,
                                                                    Direction = ParameterDirection.Input
                                                                };

                        ///////////////////  SetParameter OPT  ////////////////////////////////////////////
                        SqlParameter OPT = new SqlParameter
                                                                 {
                                                                     ParameterName = "@OPT",
                                                                     SqlDbType = SqlDbType.NVarChar,
                                                                     Value = data.OPT,
                                                                     Direction = ParameterDirection.Input
                                                                 };

                        ///////////////////  SetParameter JIKU  ////////////////////////////////////////////
                        SqlParameter JIKU = new SqlParameter
                                                                 {
                                                                     ParameterName = "@JIKU",
                                                                     SqlDbType = SqlDbType.NVarChar,
                                                                     Value = data.JIKU,
                                                                     Direction = ParameterDirection.Input
                                                                 };

                        ///////////////////  SetParameter FILLER2  ////////////////////////////////////////////
                        SqlParameter FILLER2 = new SqlParameter
                                                                {
                                                                    ParameterName = "@FILLER2",
                                                                    SqlDbType = SqlDbType.NVarChar,
                                                                    Value = data.FILLER2,
                                                                    Direction = ParameterDirection.Input
                                                                };

                        ///////////////////  SetParameter DAI  ////////////////////////////////////////////
                        SqlParameter DAI = new SqlParameter
                                                                 {
                                                                     ParameterName = "@DAI",
                                                                     SqlDbType = SqlDbType.Int,
                                                                     Value = data.DAI,
                                                                     Direction = ParameterDirection.Input
                                                                 };

                        ///////////////////  SetParameter MC  ////////////////////////////////////////////
                        SqlParameter MC = new SqlParameter
                                                                 {
                                                                     ParameterName = "@MC",
                                                                     SqlDbType = SqlDbType.NVarChar,
                                                                     Value = data.MC,
                                                                     Direction = ParameterDirection.Input
                                                                 };

                        ///////////////////  SetParameter SIMUKE  ////////////////////////////////////////////
                        SqlParameter SIMUKE = new SqlParameter
                                                                 {
                                                                     ParameterName = "@SIMUKE",
                                                                     SqlDbType = SqlDbType.NVarChar,
                                                                     Value = data.SIMUKE,
                                                                     Direction = ParameterDirection.Input
                                                                 };

                        ///////////////////  SetParameter E0  ////////////////////////////////////////////
                        SqlParameter E0 = new SqlParameter
                                                                 {
                                                                     ParameterName = "@E0",
                                                                     SqlDbType = SqlDbType.NVarChar,
                                                                     Value = data.E0,
                                                                     Direction = ParameterDirection.Input
                                                                 };

                        ///////////////////  SetParameter BUBAN  ////////////////////////////////////////////
                        SqlParameter BUBAN = new SqlParameter
                                                                 {
                                                                     ParameterName = "@BUBAN",
                                                                     SqlDbType = SqlDbType.NVarChar,
                                                                     Value = data.BUBAN,
                                                                     Direction = ParameterDirection.Input
                                                                 };

                        ///////////////////  SetParameter TANTO  ////////////////////////////////////////////
                        SqlParameter TANTO = new SqlParameter
                                                                 {
                                                                     ParameterName = "@TANTO",
                                                                     SqlDbType = SqlDbType.NVarChar,
                                                                     Value = data.TANTO,
                                                                     Direction = ParameterDirection.Input
                                                                 };

                        ///////////////////  SetParameter GR  ////////////////////////////////////////////
                        SqlParameter GR = new SqlParameter
                                                                 {
                                                                     ParameterName = "@GR",
                                                                     SqlDbType = SqlDbType.NVarChar,
                                                                     Value = data.GR,
                                                                     Direction = ParameterDirection.Input
                                                                 };

                        ///////////////////  SetParameter KIGO  ////////////////////////////////////////////
                        SqlParameter KIGO = new SqlParameter
                                                                 {
                                                                     ParameterName = "@KIGO",
                                                                     SqlDbType = SqlDbType.NVarChar,
                                                                     Value = data.KIGO,
                                                                     Direction = ParameterDirection.Input
                                                                 };

                        ///////////////////  SetParameter MAKR  ////////////////////////////////////////////
                        SqlParameter MAKR = new SqlParameter
                                                                 {
                                                                     ParameterName = "@MAKR",
                                                                     SqlDbType = SqlDbType.NVarChar,
                                                                     Value = data.MAKR,
                                                                     Direction = ParameterDirection.Input
                                                                 };

                        ///////////////////  SetParameter KOSUU  ////////////////////////////////////////////
                        SqlParameter KOSUU = new SqlParameter
                                                                 {
                                                                     ParameterName = "@KOSUU",
                                                                     SqlDbType = SqlDbType.Int,
                                                                     Value = data.KOSUU,
                                                                     Direction = ParameterDirection.Input
                                                                 };

                        ///////////////////  SetParameter KISYU  ////////////////////////////////////////////
                        SqlParameter KISYU = new SqlParameter
                                                                 {
                                                                     ParameterName = "@KISYU",
                                                                     SqlDbType = SqlDbType.NVarChar,
                                                                     Value = data.KISYU,
                                                                     Direction = ParameterDirection.Input
                                                                 };

                        ///////////////////  SetParameter MEWISYO  ////////////////////////////////////////////
                        SqlParameter MEWISYO = new SqlParameter
                                                                 {
                                                                     ParameterName = "@MEWISYO",
                                                                     SqlDbType = SqlDbType.NVarChar,
                                                                     Value = data.MEWISYO,
                                                                     Direction = ParameterDirection.Input
                                                                 };

                        ///////////////////  SetParameter FYMD  ////////////////////////////////////////////
                        SqlParameter FYMD = new SqlParameter
                                                                 {
                                                                     ParameterName = "@FYMD",
                                                                     SqlDbType = SqlDbType.DateTime,
                                                                     Value = data.FYMD,
                                                                     Direction = ParameterDirection.Input
                                                                 };

                        ///////////////////  SetParameter SEIHINCD  ////////////////////////////////////////////
                        SqlParameter SEIHINCD = new SqlParameter
                                                                 {
                                                                     ParameterName = "@SEIHINCD",
                                                                     SqlDbType = SqlDbType.NVarChar,
                                                                     Value = data.SEIHINCD,
                                                                     Direction = ParameterDirection.Input
                                                                 };

                        ///////////////////  SetParameter SEHINJNO  ////////////////////////////////////////////
                        SqlParameter SEHINJNO = new SqlParameter
                                                                 {
                                                                     ParameterName = "@SEHINJNO",
                                                                     SqlDbType = SqlDbType.NVarChar,
                                                                     Value = data.SEHINJNO,
                                                                     Direction = ParameterDirection.Input
                                                                 };

                       

                        ///////////////////  SetParameter FileName  ////////////////////////////////////////////
                        SqlParameter FileName = new SqlParameter
                                                                {
                                                                    ParameterName = "@FileName",
                                                                    SqlDbType = SqlDbType.NVarChar,
                                                                    Value = data.FileName,
                                                                    Direction = ParameterDirection.Input
                                                                };

                        ///////////////////  SetParameter LineNumber  ////////////////////////////////////////////
                        SqlParameter LineNumber = new SqlParameter
                                                                {
                                                                    ParameterName = "@LineNumber",
                                                                    SqlDbType = SqlDbType.Int,
                                                                    Value = data.LineNumber,
                                                                    Direction = ParameterDirection.Input
                                                                };

                        ///////////////////  SetParameter LineNumber  ////////////////////////////////////////////   Position = Position++
                        position = position + 1;
                        SqlParameter Position = new SqlParameter
                                                                {
                                                                    ParameterName = "@Position",
                                                                    SqlDbType = SqlDbType.Int,
                                                                    Value = position,
                                                                    Direction = ParameterDirection.Input
                                                                };

                        ///////////////////  SetParameter CreateBy  //////////////////////////////////////////////////
                        SqlParameter CreateBy = new SqlParameter
                                                                {
                                                                    ParameterName = "@CreateBy",
                                                                    SqlDbType = SqlDbType.NVarChar,
                                                                    Value = UserName,
                                                                    Direction = ParameterDirection.Input
                                                                };
                        /////////////////// Create DateTime  //////////////////////////////////////////////////
                        SqlParameter CurrentDate = new SqlParameter
                                                                    {
                                                                        ParameterName = "@CurrentDate",
                                                                        SqlDbType = SqlDbType.DateTime,
                                                                        Value = insertDateTime,
                                                                        Direction = ParameterDirection.Input
                                                                    };

                        cmd.Parameters.Add(WAYMD);
                        cmd.Parameters.Add(SEQ);
                        cmd.Parameters.Add(KATASIKI);
                        cmd.Parameters.Add(MEISHO);
                        cmd.Parameters.Add(FILLER1);
                        cmd.Parameters.Add(OPT);
                        cmd.Parameters.Add(JIKU);
                        cmd.Parameters.Add(FILLER2);
                        cmd.Parameters.Add(DAI);
                        cmd.Parameters.Add(MC);
                        cmd.Parameters.Add(SIMUKE);
                        cmd.Parameters.Add(E0);
                        cmd.Parameters.Add(BUBAN);
                        cmd.Parameters.Add(TANTO);
                        cmd.Parameters.Add(GR);
                        cmd.Parameters.Add(KIGO);
                        cmd.Parameters.Add(MAKR);
                        cmd.Parameters.Add(KOSUU);
                        cmd.Parameters.Add(KISYU);
                        cmd.Parameters.Add(MEWISYO);
                        cmd.Parameters.Add(FYMD);
                        cmd.Parameters.Add(SEIHINCD);
                        cmd.Parameters.Add(SEHINJNO);
                        cmd.Parameters.Add(FileName);
                        cmd.Parameters.Add(LineNumber);
                        cmd.Parameters.Add(Position);
                        cmd.Parameters.Add(CreateBy);
                        cmd.Parameters.Add(CurrentDate);

                        int row = cmd.ExecuteNonQuery();
                        affectedRows = affectedRows + row;
                    }
                    /////////////////////////////////////////////////////////////  BU_Mastarに対して　////////////////////////////////////////////////////
                    /////////////////////////////////////////////////////////////  SetParameter設定  /////////////////////////////////////////////////////
                    SqlCommand cmd1 = new SqlCommand()
                                                        {
                                                            CommandText = ImportDataDao.SP_BU_Mastar_SelectInsertUpdateDelete(),
                                                            Connection = connection,
                                                            CommandType = CommandType.Text,
                                                            Transaction = transaction
                                                        };
                    var statementType = "Insert";
                    foreach (var buBan in buMastars)
                    {
                        //パラメータ初期化
                        cmd1.Parameters.Clear();
                        SqlParameter User = new SqlParameter
                                                                {
                                                                    ParameterName = "@User",
                                                                    SqlDbType = SqlDbType.NVarChar,
                                                                    Value = UserName,
                                                                    Direction = ParameterDirection.Input
                                                                };
                        SqlParameter StatementType = new SqlParameter
                                                                {
                                                                    ParameterName = "@StatementType",
                                                                    SqlDbType = SqlDbType.NVarChar,
                                                                    Value = statementType,
                                                                    Direction = ParameterDirection.Input
                                                                };

                        SqlParameter BUBAN = new SqlParameter
                                                                {
                                                                    ParameterName = "@BUBAN",
                                                                    SqlDbType = SqlDbType.NVarChar,
                                                                    Value = buBan.BUBAN,
                                                                    Direction = ParameterDirection.Input
                                                                };
                        SqlParameter KIGO = new SqlParameter
                                                                {
                                                                    ParameterName = "@KIGO",
                                                                    SqlDbType = SqlDbType.NVarChar,
                                                                    Value = buBan.KIGO,
                                                                    Direction = ParameterDirection.Input
                                                                };
                        SqlParameter MEWISYO = new SqlParameter
                                                                {
                                                                    ParameterName = "@MEWISYO",
                                                                    SqlDbType = SqlDbType.NVarChar,
                                                                    Value = buBan.MEWISYO,
                                                                    Direction = ParameterDirection.Input
                                                                };
                        SqlParameter Nyusu = new SqlParameter
                                                                {
                                                                    ParameterName = "@Nyusu",
                                                                    SqlDbType = SqlDbType.Int,
                                                                    Value = buBan.Nyusu,
                                                                    Direction = ParameterDirection.Input
                                                                };
                        SqlParameter KatakanaName = new SqlParameter
                                                                {
                                                                    ParameterName = "@KatakanaName",
                                                                    SqlDbType = SqlDbType.NVarChar,
                                                                    Value = buBan.KatakanaName,
                                                                    Direction = ParameterDirection.Input
                                                                };

                        SqlParameter StausCode = new SqlParameter
                                                                {
                                                                    ParameterName = "@StausCode",
                                                                    SqlDbType = SqlDbType.Int,
                                                                    Direction = System.Data.ParameterDirection.Output
                                                                };

                        cmd1.Parameters.Add(User);
                        cmd1.Parameters.Add(StatementType);
                        cmd1.Parameters.Add(BUBAN);
                        cmd1.Parameters.Add(KIGO);
                        cmd1.Parameters.Add(MEWISYO);
                        cmd1.Parameters.Add(Nyusu);
                        cmd1.Parameters.Add(KatakanaName);
                        cmd1.Parameters.Add(StausCode);

                        int row = cmd1.ExecuteNonQuery();
                    }
/////////////////////////////////////////////////////////////  [File_Import_Log]に対して  /////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////  SetParameter設定  ////////////////////////////////////////////////////////////
                    SqlCommand cmd2 = new SqlCommand()
                                                        {
                                                            CommandText = ImportDataDao.SP_File_Import_Log_Insert(),
                                                            Connection = connection,
                                                            CommandType = CommandType.Text,
                                                            Transaction = transaction
                                                        };
                    statementType = "Insert";
                    foreach (var file in listFileName)
                    {
                        var maxLineNumber = (from item in datas
                                    where (item.FileName == file)
                                    select item).ToList().Max(x => x.LineNumber);
                        var maxPosition = (from item in datas
                                             where (item.FileName == file)
                                             select item).ToList().Max(x => x.Position);
                        //パラメータ初期化
                        cmd2.Parameters.Clear();
                        //SqlParameter StatementType = new SqlParameter
                        //                                                {
                        //                                                    ParameterName = "@StatementType",
                        //                                                    SqlDbType = SqlDbType.NVarChar,
                        //                                                    Value = statementType,
                        //                                                    Direction = ParameterDirection.Input
                        //                                                };
                        SqlParameter User = new SqlParameter
                                                                        {
                                                                            ParameterName = "@User",
                                                                            SqlDbType = SqlDbType.NVarChar,
                                                                            Value = UserName,
                                                                            Direction = ParameterDirection.Input
                                                                        };
                        SqlParameter FileName = new SqlParameter
                                                                        {
                                                                            ParameterName = "@FileName",
                                                                            SqlDbType = SqlDbType.NVarChar,
                                                                            Value = file,
                                                                            Direction = ParameterDirection.Input
                                                                        };
                        SqlParameter TotalLine = new SqlParameter
                                                                        {
                                                                            ParameterName = "@TotalLine",
                                                                            SqlDbType = SqlDbType.Int,
                                                                            Value = maxLineNumber,
                                                                            Direction = ParameterDirection.Input
                                                                        };
                        SqlParameter MaxPosition = new SqlParameter
                                                                        {
                                                                            ParameterName = "@MaxPosition",
                                                                            SqlDbType = SqlDbType.Int,
                                                                            Value = maxPosition,
                                                                            Direction = ParameterDirection.Input
                                                                        };
                        SqlParameter StausCode = new SqlParameter
                                                                        {
                                                                            ParameterName = "@StausCode",
                                                                            SqlDbType = SqlDbType.Int,
                                                                            Direction = ParameterDirection.Output
                                                                        };

                        //cmd2.Parameters.Add(StatementType);
                        cmd2.Parameters.Add(User);
                        cmd2.Parameters.Add(FileName);
                        cmd2.Parameters.Add(TotalLine);
                        cmd2.Parameters.Add(MaxPosition);
                        cmd2.Parameters.Add(StausCode);

                        int row = cmd2.ExecuteNonQuery();
                    }

                    transaction.Commit();
                    respon.Status = "OK";
                    respon.Resmess = $@"ファイルのデータがデータベースに保存されました
                                     (追加: {affectedRows}件)";
                }
                catch(Exception ex)
                {
                    var error = ex.Message;
                    transaction.Rollback();
                    respon.Status = "NG";
                    respon.Resmess = $@"ファイル読み込み中にエラーが発生しましたのでデータがデータベースに保存されていません! | Error LineNo : {lineNo} (追加: 0件)
                                     Error Message : {error}";

                    return respon;
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
                return respon;
            }
        }

        public Tuple<List<string>, List<string>> FindDropList(DateTime date)
        {
            return ImportDataDao.GetDropListItems(date);
        }

        public List<DataModel> FindDataOfLastTime(ImportViewModel viewModel)
        {
            try
            {
                return ImportDataDao.GetDataOfLastTime(viewModel);
            }
            catch
            {
                throw;
            }
        }

        public List<DataModel> FindDataOfLastTimeInit(DateTime date)
        {
            try
            {
                return ImportDataDao.GetDataOfLastTimeInit(date);
            }
            catch
            {
                throw;
            }
        }

        public int DeleteData(DateTime date)
        {
            try
            {
                return ImportDataDao.DeleteDataOnToday(date);
            }
            catch
            {
                throw;
            }
        }
    }
}
