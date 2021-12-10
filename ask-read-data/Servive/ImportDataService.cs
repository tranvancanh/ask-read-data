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

namespace ask_read_data.Servive
{
    public class ImportDataService : IImportData
    {
        public ResponResult ImportDataDB(List<object> datas1, List<Claim> Claims)
        {
            List<DataModel> datas = new List<DataModel>();
            foreach (var data in datas1)
            {
                DataModel item = (DataModel)data;
                datas.Add(item);
            }

            var respon = new ResponResult();
            int affectedRows = 0;
            var ConnectionString = new GetConnectString().ConnectionString;
            using (var connection = new SqlConnection(ConnectionString))
            {
                int lineNo = 0;
                SqlTransaction transaction = null;
                 
                connection.Open();
                transaction = connection.BeginTransaction();
                //Create the command object
                SqlCommand cmd = new SqlCommand()
                                                    {
                                                        CommandText = "SP_DataImportInsert",
                                                        Connection = connection,
                                                        CommandType = CommandType.StoredProcedure,
                                                        Transaction = transaction
                                                    };
                try 
                { 
                    foreach(var data in datas)
                    {
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
                        SqlParameter Position = new SqlParameter
                        {
                            ParameterName = "@Position",
                            SqlDbType = SqlDbType.Int,
                            Value = data.Position,
                            Direction = ParameterDirection.Input
                        };

                        ///////////////////  SetParameter CreateBy  ////////////////////////////////////////////
                        SqlParameter CreateBy = new SqlParameter
                                                                {
                                                                    ParameterName = "@CreateBy",
                                                                    SqlDbType = SqlDbType.NVarChar,
                                                                    Value = Claims.Where(c => c.Type == ClaimTypes.Name).First().Value,
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

                        int row = cmd.ExecuteNonQuery();
                        affectedRows = affectedRows + row;
                    }

                    transaction.Commit();
                    respon.Status = "OK";
                    respon.Resmess = $@"ファイルのデータがデータベースに保存されました(追加: {affectedRows}件)";
                }
                catch(Exception ex)
                {
                    var error = ex.Message;
                    transaction.Rollback();
                    respon.Status = "NG";
                    respon.Resmess = $@"ファイル読み込み中にエラーが発生しましたのでデータがデータベースに保存されていません! | Error LineNo : {lineNo} (追加: 0件)";

                    return respon;
                }
                finally
                {
                    //close connection
                    connection.Close();
                    // connection dispose
                    connection.Dispose();
                }
                return respon;
            }
        }
    }
}
