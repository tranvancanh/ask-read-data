﻿using ask_read_data.Controllers;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace mclogi.common
{
    /// <summary>
    /// ユーティリティクラス
    /// </summary>
    public class Utility
    {
        #region const

        /// <summary>セッション情報キー：ユーザーコード</summary>
        public const string SESSION_KEY_USERCD = "UserCd";
        /// <summary>セッション情報キー：ユーザー名</summary>
        public const string SESSION_KEY_USERNAME = "UserName";
        /// <summary>セッション情報キー：システムレベル</summary>
        public const string SESSION_KEY_SYSTEMLEVEL = "SystemLevel";
        /// <summary>拡張子：xlsx</summary>
        public const string EXTENSION_XLSX = ".xlsx";
        /// <summary>拡張子：csv</summary>
        public const string EXTENSION_CSV = ".csv";

        #endregion

        #region NullToBlank
        /// <summary>
        /// Nullを空文字に変換する
        /// </summary>
        /// <param name="value">変換元の値</param>
        /// <returns>変換後の値</returns>
        public static string NullToBlank(object value)
        {
            // NULL、DBNullのときは空文字に変換する
            if (value == null || value == DBNull.Value)
            {
                return string.Empty;
            }
            return Convert.ToString(value);
        }
        #endregion

        #region ConvertDateTime
        /// <summary>
        /// 日付変換する。DBNullはnullにする
        /// </summary>
        /// <param name="value">変換元の値</param>
        /// <returns>変換後の値</returns>
        public static DateTime? ConvertDateTime(object value)
        {
            if (value == null || value == DBNull.Value)
            {
                return null;
            }
            try
            {
                return Convert.ToDateTime(value);
            }
            catch (FormatException)
            {
                return null;
            }
        }
        #endregion

        #region ChangeMondy
        /// <summary>
        /// 月曜日の日付に変換する
        /// </summary>
        public DateTime ChangeMonday(DateTime date)
        {
            switch (date.DayOfWeek.ToString("d"))
            {
                case "0":
                    date = date.AddDays(-6);
                    break;
                case "1":
                    break;
                case "2":
                    date = date.AddDays(-1);
                    break;
                case "3":
                    date = date.AddDays(-2);
                    break;
                case "4":
                    date = date.AddDays(-3);
                    break;
                case "5":
                    date = date.AddDays(-4);
                    break;
                case "6":
                    date = date.AddDays(-5);
                    break;
                default:
                    break;
            }
            return date;
        }
        #endregion

        #region DateTimeTryParse
        /// <summary>
        /// DateTime型に変換可能かboolで返す
        /// </summary>
        /// <param name="value">変換元の値</param>
        /// <returns>変換後の値</returns>
        public bool DateTimeTryParse(string value)
        {
            if (DateTime.TryParse(value, out DateTime dateTime))
            {
                // DateTime型に変換で来たらtrueを返す
                return true;
            }
            else
            {
                return false;
            }
        }
        #endregion

        #region Excel読込
        /// <summary>
        /// Excel読込
        /// </summary>
        /// <param name="fileno">ファイル番号</param>
        /// <param name="filePath">読込ファイルフルパス</param>
        /// <param name="columNameList">列名リスト</param>
        /// <param name="errMessage">エラーメッセージ</param>
        /// <param name="startRowIndex">データ開始行番号(初期値：2)</param>
        /// <param name="titleRowCount">タイトル行数(初期値：1)</param>
        /// <param name="titleCheckCount">タイトルのチェック行(初期値：1)</param>
        /// <param name="requiredColumnNo">必須データ列番号(初期値：1)</param>
        /// <returns></returns>
        public DataTable ImportExcel(string fileno, string filePath, List<string> columNameList, int startRowIndex, int titleRowCount, int titleCheckCount, int requiredColumnNo, out string errMessage)
        {
            errMessage = string.Empty;

            // 列名指定リストが未設定の時は中断
            if (columNameList == null && columNameList.Count == 0)
            {
                return null;
            }
            // ファイル不在の場合は中断
            if (!File.Exists(filePath))
            {
                errMessage = "Upload failed.";
                return null;
            }

            // 拡張子が「xlsx」でないときは中断
            if (!EXTENSION_XLSX.Equals(Path.GetExtension(filePath)))
            {
                errMessage = "The selected file is incorrect.";
                return null;
            }

            DataTable resultDt = null;

            try
            {
                FileInfo fileInfo = new FileInfo(filePath);

                using (var package = new ExcelPackage(fileInfo))
                {
                    // 1シート目を取得
                    using (ExcelWorksheet sheet = package.Workbook.Worksheets[1])
                    {
                        // データがないときは中断
                        if (sheet.Dimension == null)
                        {
                            errMessage = "No data.";
                            return null;
                        }

                        //// データのある列数と、列名指定リストの件数が異なる場合は中断
                        //if (sheet.Dimension.Columns != columNameList.Count)
                        //{
                        //    errMessage = "The contents of the file are incorrect.";
                        //    return null;
                        //}

                        // データのある行数が、タイトルの行数以下の場合は中断
                        if (sheet.Dimension.Rows <= titleRowCount)
                        {
                            errMessage = "No data.";
                            return null;
                        }

                        // データのある列数と、列名指定リストの件数が異なる場合は中断
                        var columnCount = 0;
                        var titleData = Utility.NullToBlank(sheet.Cells[titleCheckCount, columnCount + 1].Value).Trim().Replace("\n", "");
                        while (!String.IsNullOrEmpty(titleData))
                        {
                            columnCount++;
                            titleData = Utility.NullToBlank(sheet.Cells[titleCheckCount, columnCount + 1].Value).Trim().Replace("\n", "");
                        }

                        //if (columnCount < columNameList.Count)
                        //{
                        //    //指定列よりも少ない場合はエラー
                        //    errMessage = "The contents of the file are incorrect.";
                        //    return null;
                        //}

                        // データテーブルのインスタンスを生成し、列追加
                        resultDt = new DataTable();
                        // 列名指定リストの分だけ列を追加
                        for (int i = 0; i < columNameList.Count; i++)
                        {
                            resultDt.Columns.Add(columNameList[i]);
                        }
                        if (fileno == "11")
                        {
                            resultDt.Columns.Add("MaterialCode_head2", Type.GetType("System.String"));
                            resultDt.Columns.Add("MaterialCode_head4", Type.GetType("System.String"));
                            resultDt.Columns.Add("MaterialCode_foot4", Type.GetType("System.String"));
                        }


                        // データ読込(開始行番号から、必須データ列の値がなくなるまで)
                        int r = startRowIndex;
                        // １行目の必須列の値を取得
                        var requiredData = Utility.NullToBlank(sheet.Cells[r, requiredColumnNo].Value).Trim().Replace("\n", "");
                        // requiredData が IsNullOrEmpty ではなくても、[U/N/D]が"END"だったら読み込みを終了するフラグ
                        var isEnd = false;

                        while (!String.IsNullOrEmpty(requiredData))
                        {
                            var newRow = resultDt.NewRow();
                            var tblColIndex = 0;

                            for (int c = 1; c <= columNameList.Count; c++)
                            {
                                newRow[tblColIndex] = Utility.NullToBlank(sheet.Cells[r, c].Value).Trim().Replace("\n", "");

                                string moji = newRow[tblColIndex].ToString();

                                if (c == 1 && moji == "END")
                                {
                                    isEnd = true;
                                    break;
                                }

                                //if (moji.Length == 10)
                                //{
                                //    if (DateTime.TryParse(@$"{moji.Substring(6, 4)}/{moji.Substring(3, 2)}/{moji.Substring(0, 2)}", out DateTime date))
                                //    {
                                //        newRow[tblColIndex] = date;
                                //    }

                                //}

                                if (columNameList[c - 1].Contains("date", StringComparison.OrdinalIgnoreCase)) // 大文字・小文字を問わない
                                {
                                    if (moji == "")
                                    {
                                        // date型 & NOT NULL なのでSQLserverの空文字変換値を入れておく
                                        newRow[tblColIndex] = "1900/01/01";
                                    }
                                    else
                                    {
                                        // dd.MM.yyyy の形式の場合は yyyy/MM/dd に変換
                                        if (moji.Length == 10 && moji.Contains("."))
                                        {
                                            moji = @$"{moji.Substring(6, 4)}/{moji.Substring(3, 2)}/{moji.Substring(0, 2)}";
                                        }

                                        // 値がDateTimeに変更できるか
                                        if (DateTimeTryParse(moji))
                                        {
                                            newRow[tblColIndex] = ConvertDateTime(moji);
                                        }
                                        else if (moji.All(char.IsDigit))
                                        {
                                            // Excelデータにyyyy-MM-dd形式で表示されている場合に、シリアル値で取得されてしまう対応
                                            var doubleMoji = double.Parse(moji);
                                            try
                                            {
                                                newRow[tblColIndex] = DateTime.FromOADate(doubleMoji);
                                            }
                                            catch (ArgumentException ex) // DateTimeに変換できない場合のエラー
                                            {
                                                errMessage = "Error: " + r + "line DateTime Format Error";
                                                return null;
                                            }
                                        }
                                        else
                                        {
                                            // 想定外のDateTime型に変更不可の場合はここを通る
                                            Console.WriteLine(moji);

                                            errMessage = "Error: " + r + "line DateTime Format Error";
                                            return null;
                                        }
                                    }
                                }

                                // int型チェック
                                if (columNameList[c - 1] == "InquiryQty")
                                {
                                    if (moji == "")
                                    {
                                        // int型 & NOT NULL なので空欄の場合は0を入れておく
                                        newRow[tblColIndex] = 0;
                                    }
                                    else if (!moji.All(char.IsDigit))
                                    {
                                        // 数値以外の値が入っていたらエラー
                                        errMessage = "Error: " + r + "line [" + columNameList[c - 1] + "] Not a number.";
                                        return null;
                                    }
                                }

                                //if (moji.Substring(2,1)=="." && moji.Substring(5, 1) == "." && moji.Length == 10)
                                //{

                                //}

                                tblColIndex++;
                            }

                            if (isEnd)
                            {
                                break;
                            }

                            if (fileno == "11")
                            {
                                newRow["MaterialCode_head2"] = newRow["ItemNo"].ToString().Substring(0, 2);
                                newRow["MaterialCode_head4"] = newRow["ItemNo"].ToString().Substring(0, 4);
                                newRow["MaterialCode_foot4"] = Microsoft.VisualBasic.Strings.Right(newRow["ItemNo"].ToString(), 4);
                            }

                            //if (fileno == "01" || fileno == "03")
                            //{
                            //    string moji2 = newRow["Type"].ToString();
                            //    if (moji2 == "Local")
                            //    {
                            //        moji2 = "L";
                            //    }
                            //    if (moji2 == "Import")
                            //    {
                            //        moji2 = "I";
                            //    }
                            //    newRow["Type"] = moji2;

                            //}

                            resultDt.Rows.Add(newRow);

                            // 行をカウントアップ
                            r++;
                            // 必須列の値を取得
                            requiredData = Utility.NullToBlank(sheet.Cells[r, requiredColumnNo].Value).Trim().Replace("\n", "");
                        }

                        //for (int r = startRowIndex; r <= sheet.Dimension.End.Row; r++)
                        //{
                        //    var newRow = resultDt.NewRow();
                        //    var tblColIndex = 0;

                        //    for (int c = 1; c <= columNameList.Count; c++)
                        //    {
                        //        newRow[tblColIndex] = Utility.NullToBlank(sheet.Cells[r, c].Value).Trim().Replace("\n", "");
                        //        tblColIndex++;
                        //    }

                        //    resultDt.Rows.Add(newRow);
                        //}

                        resultDt.AcceptChanges();
                    }
                }
            }
            catch (Exception ex)
            {
                // 読み込み失敗
                errMessage = "Upload failed.";
                resultDt = null;
            }

            return resultDt;
        }
        #endregion

        #region Excel読込
        /// <summary>
        /// Excel読込
        /// </summary>
        /// <param name="fileno">ファイル番号</param>
        /// <param name="filePath">読込ファイルフルパス</param>
        /// <param name="columNameList">列名リスト</param>
        /// <param name="errMessage">エラーメッセージ</param>
        /// <param name="startRowIndex">データ開始行番号(初期値：2)</param>
        /// <param name="titleRowCount">タイトル行数(初期値：1)</param>
        /// <param name="titleCheckCount">タイトルのチェック行(初期値：1)</param>
        /// <param name="requiredColumnNo">必須データ列番号(初期値：1)</param>
        /// <returns></returns>
        public DataTable ImportMasterExcel(string fileno, string filePath, List<string> columNameList, int startRowIndex, int titleRowCount, int titleCheckCount, int requiredColumnNo, out string errMessage)
        {
            errMessage = string.Empty;

            // 列名指定リストが未設定の時は中断
            if (columNameList == null && columNameList.Count == 0)
            {
                return null;
            }
            // ファイル不在の場合は中断
            if (!File.Exists(filePath))
            {
                errMessage = "Upload failed.";
                return null;
            }

            // 拡張子が「xlsx」でないときは中断
            if (!EXTENSION_XLSX.Equals(Path.GetExtension(filePath)))
            {
                errMessage = "The selected file is incorrect.";
                return null;
            }

            DataTable resultDt = null;

            try
            {
                FileInfo fileInfo = new FileInfo(filePath);

                using (var package = new ExcelPackage(fileInfo))
                {
                    // 1シート目を取得
                    using (ExcelWorksheet sheet = package.Workbook.Worksheets[1])
                    {
                        // データがないときは中断
                        if (sheet.Dimension == null)
                        {
                            errMessage = "No data.";
                            return null;
                        }

                        //// データのある列数と、列名指定リストの件数が異なる場合は中断
                        //if (sheet.Dimension.Columns != columNameList.Count)
                        //{
                        //    errMessage = "The contents of the file are incorrect.";
                        //    return null;
                        //}

                        // データのある行数が、タイトルの行数以下の場合は中断
                        if (sheet.Dimension.Rows <= titleRowCount)
                        {
                            errMessage = "No data.";
                            return null;
                        }

                        // データのある列数と、列名指定リストの件数が異なる場合は中断
                        var columnCount = 0;
                        var titleData = Utility.NullToBlank(sheet.Cells[titleCheckCount, columnCount + 1].Value).Trim().Replace("\n", "");
                        while (!String.IsNullOrEmpty(titleData))
                        {
                            columnCount++;
                            titleData = Utility.NullToBlank(sheet.Cells[titleCheckCount, columnCount + 1].Value).Trim().Replace("\n", "");
                        }

                        //if (columnCount < columNameList.Count)
                        //{
                        //    //指定列よりも少ない場合はエラー
                        //    errMessage = "The contents of the file are incorrect.";
                        //    return null;
                        //}

                        // データテーブルのインスタンスを生成し、列追加
                        resultDt = new DataTable();
                        // 列名指定リストの分だけ列を追加
                        for (int i = 0; i < columNameList.Count; i++)
                        {
                            resultDt.Columns.Add(columNameList[i]);
                        }

                        // データ読込(開始行番号から、必須データ列の値がなくなるまで)
                        int r = startRowIndex;
                        // １行目の必須列の値を取得
                        var requiredData = Utility.NullToBlank(sheet.Cells[r, requiredColumnNo].Value).Trim().Replace("\n", "");
                        // requiredData が IsNullOrEmpty ではなくても、[U/N/D]が"END"だったら読み込みを終了するフラグ
                        var isEnd = false;

                        while (!String.IsNullOrEmpty(requiredData))
                        {
                            var newRow = resultDt.NewRow();
                            var tblColIndex = 0;

                            for (int c = 1; c <= columNameList.Count; c++)
                            {
                                newRow[tblColIndex] = Utility.NullToBlank(sheet.Cells[r, c].Value).Trim().Replace("\n", "");

                                string moji = newRow[tblColIndex].ToString();

                                // 畠山テスト　"DBAR" -が0になる　←とりあえずこのままで
                                if (columNameList[c - 1] == "DBAR")
                                {
                                    newRow[tblColIndex] = Utility.NullToBlank(sheet.Cells[r, c].Value).Trim().Replace("\n", "");
                                    var value = newRow[tblColIndex];
                                }

                                // 1列目チェック
                                if (c == 1)
                                {
                                    if (moji == "END")
                                    {
                                        isEnd = true;
                                        break;
                                    }
                                    else if (moji != "" && moji != "U" && moji != "N" && moji != "D")
                                    {
                                        errMessage = "Error: " + r + "line [U/N/D] Unrecognizable";
                                        return null;
                                    }
                                }

                                // DateTime型チェック
                                if (columNameList[c - 1].Contains("date", StringComparison.OrdinalIgnoreCase)) // 大文字・小文字を問わない
                                {
                                    if (moji == "")
                                    {
                                        // date型 & NOT NULL なのでSQLserverの空文字変換値を入れておく
                                        newRow[tblColIndex] = "1900/01/01";
                                    }
                                    else
                                    {
                                        // dd.MM.yyyy の形式の場合は yyyy/MM/dd に変換
                                        if (moji.Length == 10 && moji.Contains("."))
                                        {
                                            moji = @$"{moji.Substring(6, 4)}/{moji.Substring(3, 2)}/{moji.Substring(0, 2)}";
                                        }

                                        // DateTimeに変更できるか
                                        if (DateTimeTryParse(moji))
                                        {
                                            newRow[tblColIndex] = ConvertDateTime(moji);
                                        }
                                        else if (moji.All(char.IsDigit))
                                        {
                                            // Excelデータにyyyy-MM-dd形式で表示されている場合に、シリアル値で取得されてしまう対応
                                            var doubleMoji = double.Parse(moji);
                                            try
                                            {
                                                newRow[tblColIndex] = DateTime.FromOADate(doubleMoji);
                                            }
                                            catch (ArgumentException ex) // DateTimeに変換できない場合のエラー
                                            {
                                                errMessage = "Error: " + r + "line [" + columNameList[c - 1] + "] Not in date format.";
                                                return null;
                                            }
                                        }
                                        else
                                        {
                                            // 想定外のDateTime型に変更不可の場合はここを通る
                                            Console.WriteLine(moji);

                                            errMessage = "Error: " + r + "line [" + columNameList[c - 1] + "] Not in date format.";
                                            return null;
                                        }
                                    }

                                }

                                // int型チェック
                                if (columNameList[c - 1] == "PartCode" || columNameList[c - 1] == "MOQ" || columNameList[c - 1] == "Qty")
                                {
                                    if (moji == "")
                                    {
                                        // int型 & NOT NULL なので空欄の場合は0を入れておく
                                        newRow[tblColIndex] = 0;
                                    }
                                    else if (!moji.All(char.IsDigit))
                                    {
                                        // 数値以外の値が入っていたらエラー
                                        errMessage = "Error: " + r + "line [" + columNameList[c - 1] + "] Not a number.";
                                        return null;
                                    }
                                }

                                tblColIndex++;
                            }

                            if (isEnd)
                            {
                                break;
                            }

                            if (fileno == "01")
                            {
                                // Excelデータに無い行　挿入先テーブルはNOT NULLなので空欄を入れておく
                                newRow["LastPart"] = "";
                                newRow["ActiveParts"] = "";
                            }

                            if (fileno == "03")
                            {
                                string moji2 = newRow["Remarks"].ToString();
                                if (moji2.Contains("non", StringComparison.OrdinalIgnoreCase)) //大文字小文字を区別しない
                                {
                                    moji2 = "Non Active";
                                }
                                else if (moji2.Contains("active", StringComparison.OrdinalIgnoreCase)) //大文字小文字を区別しない
                                {
                                    moji2 = "Active";
                                }
                                newRow["Remarks"] = moji2;
                            }

                            if (fileno == "01" || fileno == "03")
                            {
                                string moji2 = newRow["Type"].ToString();
                                //if (moji2 == "Local")
                                if (moji2.Contains("local", StringComparison.OrdinalIgnoreCase)) //大文字小文字を区別しない
                                {
                                    moji2 = "L";
                                }
                                //if (moji2 == "Import")
                                else if (moji2.Contains("import", StringComparison.OrdinalIgnoreCase)) //大文字小文字を区別しない
                                {
                                    moji2 = "I";
                                }
                                else if (moji2.Contains("export", StringComparison.OrdinalIgnoreCase)) //大文字小文字を区別しない
                                {
                                    moji2 = "E";
                                }
                                newRow["Type"] = moji2;
                            }


                            resultDt.Rows.Add(newRow);

                            // 行をカウントアップ
                            r++;
                            // 必須列の値を取得
                            requiredData = Utility.NullToBlank(sheet.Cells[r, requiredColumnNo].Value).Trim().Replace("\n", "");
                        }

                        //for (int r = startRowIndex; r <= sheet.Dimension.End.Row; r++)
                        //{
                        //    var newRow = resultDt.NewRow();
                        //    var tblColIndex = 0;

                        //    for (int c = 1; c <= columNameList.Count; c++)
                        //    {
                        //        newRow[tblColIndex] = Utility.NullToBlank(sheet.Cells[r, c].Value).Trim().Replace("\n", "");
                        //        tblColIndex++;
                        //    }

                        //    resultDt.Rows.Add(newRow);
                        //}

                        if (fileno == "03")
                        {
                            //// ExcelデータにPercent列がないので、実際のテーブルに合わせた位置に挿入する
                            //DataColumn dc = new DataColumn();
                            //dc.ColumnName = "Percent";
                            //dc.DefaultValue = "";
                            //resultDt.Columns.Add(dc);
                            //dc.SetOrdinal(8);

                            // ExcelデータのPriceCord列はSALES列の後ろにあるが
                            // テーブルでは最後の列なので合わせる
                            resultDt.Columns["PriceCord"].SetOrdinal(16);
                        }

                        resultDt.AcceptChanges();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                // 読み込み失敗
                errMessage = "Upload failed.";
                resultDt = null;
            }

            return resultDt;
        }
        #endregion

        #region Excel読込
        /// <summary>
        /// Excel読込
        /// </summary>
        /// <param name="fileno">ファイル番号</param>
        /// <param name="filePath">読込ファイルフルパス</param>
        /// <param name="columNameList">列名リスト</param>
        /// <param name="errMessage">エラーメッセージ</param>
        /// <param name="startRowIndex">データ開始行番号(初期値：2)</param>
        /// <param name="titleRowCount">タイトル行数(初期値：1)</param>
        /// <param name="titleCheckCount">タイトルのチェック行(初期値：1)</param>
        /// <param name="requiredColumnNo">必須データ列番号(初期値：1)</param>
        /// <returns></returns>
        public DataTable ImportDummyExcel(string fileno, string filePath, List<string> columNameList, int startRowIndex, int titleRowCount, int titleCheckCount, int requiredColumnNo, out string errMessage)
        {
            errMessage = string.Empty;

            // 列名指定リストが未設定の時は中断
            if (columNameList == null && columNameList.Count == 0)
            {
                return null;
            }
            // ファイル不在の場合は中断
            if (!File.Exists(filePath))
            {
                errMessage = "Upload failed.";
                return null;
            }

            // 拡張子が「xlsx」でないときは中断
            if (!EXTENSION_XLSX.Equals(Path.GetExtension(filePath)))
            {
                errMessage = "The selected file is incorrect.";
                return null;
            }

            DataTable resultDt = null;

            try
            {
                FileInfo fileInfo = new FileInfo(filePath);

                using (var package = new ExcelPackage(fileInfo))
                {
                    // 1シート目を取得
                    using (ExcelWorksheet sheet = package.Workbook.Worksheets[1])
                    {
                        // データがないときは中断
                        if (sheet.Dimension == null)
                        {
                            errMessage = "No data.";
                            return null;
                        }

                        //// データのある列数と、列名指定リストの件数が異なる場合は中断
                        //if (sheet.Dimension.Columns != columNameList.Count)
                        //{
                        //    errMessage = "The contents of the file are incorrect.";
                        //    return null;
                        //}

                        // データのある行数が、タイトルの行数以下の場合は中断
                        if (sheet.Dimension.Rows <= titleRowCount)
                        {
                            errMessage = "No data.";
                            return null;
                        }

                        // データのある列数と、列名指定リストの件数が異なる場合は中断
                        var columnCount = 0;
                        var titleData = Utility.NullToBlank(sheet.Cells[titleCheckCount, columnCount + 1].Value).Trim().Replace("\n", "");
                        while (!String.IsNullOrEmpty(titleData))
                        {
                            columnCount++;
                            titleData = Utility.NullToBlank(sheet.Cells[titleCheckCount, columnCount + 1].Value).Trim().Replace("\n", "");
                        }

                        //if (columnCount < columNameList.Count)
                        //{
                        //    //指定列よりも少ない場合はエラー
                        //    errMessage = "The contents of the file are incorrect.";
                        //    return null;
                        //}

                        // データテーブルのインスタンスを生成し、列追加
                        resultDt = new DataTable();
                        // 列名指定リストの分だけ列を追加
                        for (int i = 0; i < columNameList.Count; i++)
                        {
                            resultDt.Columns.Add(columNameList[i]);
                        }

                        // データ読込(開始行番号から、必須データ列の値がなくなるまで)
                        int r = startRowIndex;
                        // １行目の必須列の値を取得
                        var requiredData = Utility.NullToBlank(sheet.Cells[r, requiredColumnNo].Value).Trim().Replace("\n", "");
                        if (requiredData == "")
                        {
                            // 値がNULLの場合はエラーを返す
                            errMessage = "Error: " + r + "line [" + columNameList[requiredColumnNo - 1] + "] No value";
                            return null;
                        }
                        // requiredData が IsNullOrEmpty ではなくても、[U/N/D]が"END"だったら読み込みを終了するフラグ
                        var isEnd = false;

                        while (!String.IsNullOrEmpty(requiredData))
                        {
                            var newRow = resultDt.NewRow();
                            var tblColIndex = 0;

                            for (int c = 1; c <= columNameList.Count; c++)
                            {
                                newRow[tblColIndex] = Utility.NullToBlank(sheet.Cells[r, c].Value).Trim().Replace("\n", "");

                                string moji = newRow[tblColIndex].ToString();

                                if (c == 1)
                                {
                                    if (moji == "END")
                                    {
                                        isEnd = true;
                                        break;
                                    }
                                }

                                if (moji == "")
                                {
                                    // 値がNULLの場合はエラーを返す
                                    errMessage = "Error: " + r + "line [" + columNameList[c - 1] + "] No value";
                                    return null;
                                }

                                if (columNameList[c - 1] == "InquiryQty")
                                {
                                    // InquiryQtyが数値に変換できなかったらエラーを返す
                                    bool isIntValue = int.TryParse(moji, out int intValue);
                                    if (!isIntValue)
                                    {
                                        errMessage = "Error: " + r + "line [" + columNameList[c - 1] + "] Not a number.";
                                        return null;
                                    }
                                }

                                if (columNameList[c - 1] == "InquiryDate")
                                {
                                    // dd.MM.yyyy の形式の場合は yyyy/MM/dd に変換
                                    if (moji.Length == 10 && moji.Contains("."))
                                    {
                                        moji = @$"{moji.Substring(6, 4)}/{moji.Substring(3, 2)}/{moji.Substring(0, 2)}";
                                    }

                                    // 値がDateTimeにできるか
                                    if (DateTimeTryParse(moji))
                                    {
                                        newRow[tblColIndex] = ConvertDateTime(moji);
                                    }
                                    else if (moji.All(char.IsDigit))
                                    {
                                        // Excelデータにyyyy-MM-dd形式で表示されている場合に、シリアル値で取得されてしまう対応
                                        var doubleMoji = double.Parse(moji);
                                        try
                                        {
                                            newRow[tblColIndex] = DateTime.FromOADate(doubleMoji);
                                        }
                                        catch (ArgumentException ex) // DateTimeに変換できない場合のエラー
                                        {
                                            errMessage = "Error: " + r + "line [InquiryDate] Not in date format.";
                                            return null;
                                        }
                                    }
                                    else
                                    {
                                        // 想定外のDateTime型に変更不可の場合はここを通る
                                        Console.WriteLine(moji);

                                        errMessage = "Error: " + r + "line [InquiryDate] Not in date format.";
                                        return null;
                                    }

                                    // InquiryDateを月曜日の日付に変更する
                                    newRow[tblColIndex] = ChangeMonday(Convert.ToDateTime(newRow[tblColIndex]));
                                }

                                tblColIndex++;
                            }

                            if (isEnd)
                            {
                                break;
                            }

                            resultDt.Rows.Add(newRow);

                            // 行をカウントアップ
                            r++;
                            // 必須列の値を取得
                            requiredData = Utility.NullToBlank(sheet.Cells[r, requiredColumnNo].Value).Trim().Replace("\n", "");
                            bool isOtherValue = false;
                            if (requiredData == "")
                            {
                                // 必須列の値が入っていなくても、他の列に値が入っていればエラーを返す
                                for (int i = 1; i < columNameList.Count; ++i)
                                {
                                    var otherValue = Utility.NullToBlank(sheet.Cells[r, columNameList.Count - i].Value).Trim().Replace("\n", "");
                                    if (otherValue != "")
                                    {
                                        isOtherValue = true;
                                        break;
                                    }
                                }

                                if (isOtherValue)
                                {
                                    errMessage = "Error: " + r + "line [" + columNameList[requiredColumnNo - 1] + "] No value";
                                    return null;
                                }
                            }
                        }

                        resultDt.AcceptChanges();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                // 読み込み失敗
                errMessage = "Upload failed.";
                resultDt = null;
            }

            return resultDt;
        }
        #endregion

        #region Excel出力
        ///// <summary>
        ///// Excel出力
        ///// </summary>
        ///// <param name="dt">出力データDataTable</param>
        ///// <param name="exportfileFullPath">出力パス(フルパス)</param>
        ///// <param name="titleRows">タイトル行情報</param>
        ///// <param name="dateTimeColumns">日付指定をする列番号</param>
        ///// <param name="sheetName">シート名</param>
        public bool ExportExcel(DataTable dt1, DataTable dt2, string exportfileFullPath, string tempFileFullPath, string buBanType, List<string> titleRows, List<int> dateTimeColumns, List<string> sheetName)
        {
            // データがない時は中断
            if (dt1 == null || dt1.Rows.Count <= 0 || dt2 == null || dt2.Rows.Count <= 0)
            {
                return false;
            }
            if(sheetName.Count < 2)
            {
                return false;
            }
            // 出力ファイルパスが未指定の場合は中断する
            if (String.IsNullOrWhiteSpace(exportfileFullPath))
            {
                return false;
            }
            // 出力フォルダが存在しない場合は中断する
            if (!Directory.Exists(Path.GetDirectoryName(exportfileFullPath)))
            {
                return false;
            }
            // 既にファイルが存在している場合は削除する
            if (File.Exists(exportfileFullPath))
            {
                File.Delete(exportfileFullPath);
            }

            try
            {
                // 出力用ファイルを生成する
                //FileInfo fileInfo = new FileInfo(exportfileFullPath);
                var newFile = new FileInfo(exportfileFullPath);
                var tempFile = new FileInfo(tempFileFullPath);

                var startIndex = 3;
                var printHeader = true;
                using (var package = new ExcelPackage(newFile, tempFile))
                {
                    ///////////////////////////////////////// Sheet 1 ////////////////////////////////////////////////////////////////////////////////////////////////////////
                    var sheet = package.Workbook.Worksheets.Where(s => s.Name == sheetName[0]).FirstOrDefault();
                    if(sheet != null)
                    {
                        // シート追加
                        //package.Workbook.Worksheets.Add(sheetName[0]);
                        // シート取得
                        // // 1シートずつ取得
                        //var sheet = package.Workbook.Worksheets.Add(sheetName[0]);
                        //using (sheet = package.Workbook.Worksheets.Add(sheetName[0]))
                        //{
                        if (titleRows != null && titleRows.Count > 0)
                        {
                            for (int i = 0; i < titleRows.Count; i++)
                            {
                                sheet.Cells[1, i + 1].Value = titleRows[i];
                            }
                            // 開始行番号をセット
                            startIndex = 4;
                            // タイトル出力済なので、列名は出力しない
                            printHeader = false;
                        }
                        // not diplay header
                        printHeader = false;
                        // データセット
                        // sheet.Cells[startIndex, 1].LoadFromDataTable(dt1, printHeader);
                        LoadDataIntoTemplateExcel(ref sheet, dt1, buBanType, "sheet1");

                        // 日付指定をする列番号がある場合
                        if (dateTimeColumns != null && dateTimeColumns.Count > 0)
                        {
                            for (int i = 0; i < dateTimeColumns.Count; i++)
                            {
                                sheet.Cells[startIndex, dateTimeColumns[i], sheet.Dimension.Rows, dateTimeColumns[i]].Style.Numberformat.Format = "yyyy/MM/dd";
                            }
                        }
                        // センターアラインエクセル
                        //sheet.Cells["A" + startIndex.ToString() + ":F" + (startIndex + dt1.Rows.Count - 1).ToString()].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        //sheet.Cells["A" + startIndex.ToString() + ":F" + (startIndex + dt1.Rows.Count - 1).ToString()].AutoFitColumns();

                        //sheet.Cells["A" + startIndex.ToString() + ":F" + (startIndex + dt1.Rows.Count - 1).ToString()].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        //sheet.Cells["A" + startIndex.ToString() + ":F" + (startIndex + dt1.Rows.Count - 1).ToString()].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        //sheet.Cells["A" + startIndex.ToString() + ":F" + (startIndex + dt1.Rows.Count - 1).ToString()].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        //sheet.Cells["A" + startIndex.ToString() + ":F" + (startIndex + dt1.Rows.Count - 1).ToString()].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                        // Upload画面のとき（ファイル名最初の2文字が数字）だけ、最後の行に"END"を追加
                        var fileHeadName = newFile.Name.Substring(0, 2);
                        if (int.TryParse(fileHeadName, out int number))
                        {
                            var lastRowNo = sheet.Dimension.End.Row;
                            sheet.InsertRow(lastRowNo + 1, 1);
                            sheet.Cells[lastRowNo + 1, 1].Value = "END";
                        };

                        // 保管
                        //package.Save();
                        //}
                    }
                    else { return false; }
                    ///////////////////////////////////////// Sheet 2 ////////////////////////////////////////////////////////////////////////////////////////////////////////
                    sheet = package.Workbook.Worksheets.Where(s => s.Name == sheetName[1]).FirstOrDefault();
                    if(sheet != null)
                    {
                        //using (sheet = package.Workbook.Worksheets.Add(sheetName[1]))
                        //{
                        //startIndex = 3;
                        //sheet = package.Workbook.Worksheets.Add(sheetName[1]);
                        //sheet.Cells.Style.Font.Size = 11; //Default font size for whole sheet
                        //sheet.Cells.Style.Font.Name = "游ゴシック"; //Default Font name for whole sheet
                        //if (buBanType == ExportExcelController.FLOOR_ASSY)
                        //{
                        //    sheet.Cells[1, 1].Value = "【生産用】FLOOR ASSY";
                        //}
                        //else if (buBanType == ExportExcelController.FLAME_ASSY)
                        //{
                        //    sheet.Cells[1, 1].Value = "【生産用】FLAME ASSY";
                        //}
                        // header clor
                        ////var colFromHex = System.Drawing.FromArgb.FromHtml("#FFFF00");
                        //sheet.Cells["A" + startIndex.ToString() + ":F" + startIndex.ToString()].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        //sheet.Cells["A" + startIndex.ToString() + ":F" + startIndex.ToString()].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                        // タイトル行が指定されているときは、タイトル行をセットする
                        if (titleRows != null && titleRows.Count > 0)
                        {
                            for (int i = 0; i < titleRows.Count; i++)
                            {
                                sheet.Cells[1, i + 1].Value = titleRows[i];
                                // センターアラインエクセル
                                sheet.Cells["A" + startIndex.ToString() + ":F" + startIndex.ToString()].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                sheet.Cells["A" + startIndex.ToString() + ":F" + startIndex.ToString()].AutoFitColumns();
                            }
                            // 開始行番号をセット
                            startIndex = 4;
                            // タイトル出力済なので、列名は出力しない
                            printHeader = false;
                        }
                        // not diplay header
                        printHeader = false;
                        // データセット
                        // sheet.Cells[startIndex, 1].LoadFromDataTable(dt2, printHeader);
                        LoadDataIntoTemplateExcel(ref sheet, dt2, buBanType, "sheet2");
                        // 日付指定をする列番号がある場合
                        if (dateTimeColumns != null && dateTimeColumns.Count > 0)
                        {
                            for (int i = 0; i < dateTimeColumns.Count; i++)
                            {
                                sheet.Cells[startIndex, dateTimeColumns[i], sheet.Dimension.Rows, dateTimeColumns[i]].Style.Numberformat.Format = "yyyy/MM/dd";
                            }
                        }
                        // センターアラインエクセル
                        //sheet.Cells["A" + startIndex.ToString() + ":F" + (startIndex + dt2.Rows.Count - 1).ToString()].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        //sheet.Cells["A" + startIndex.ToString() + ":F" + (startIndex + dt2.Rows.Count - 1).ToString()].AutoFitColumns();

                        //sheet.Cells["A" + startIndex.ToString() + ":F" + (startIndex + dt2.Rows.Count - 1).ToString()].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        //sheet.Cells["A" + startIndex.ToString() + ":F" + (startIndex + dt2.Rows.Count - 1).ToString()].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        //sheet.Cells["A" + startIndex.ToString() + ":F" + (startIndex + dt2.Rows.Count - 1).ToString()].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        //sheet.Cells["A" + startIndex.ToString() + ":F" + (startIndex + dt2.Rows.Count - 1).ToString()].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        // Upload画面のとき（ファイル名最初の2文字が数字）だけ、最後の行に"END"を追加
                        var fileHeadName1 = newFile.Name.Substring(0, 2);
                        if (int.TryParse(fileHeadName1, out int number1))
                        {
                            var lastRowNo = sheet.Dimension.End.Row;
                            sheet.InsertRow(lastRowNo + 1, 1);
                            sheet.Cells[lastRowNo + 1, 1].Value = "END";
                        };
                    }
                    else { return false; }

                    // 保管
                    package.Save();
                    //}
                }

            }
            catch
            {
                // 失敗したときは出力用ファイルを削除する
                if (File.Exists(exportfileFullPath))
                {
                    File.Delete(exportfileFullPath);
                }

                throw;
            }

            return true;
        }

        #endregion

        private bool LoadDataIntoTemplateExcel(ref ExcelWorksheet worksheet, DataTable data, string buBanType, string sheetNo)
        {
            if (data.Rows.Count <= 0) { return false; }
            worksheet.Cells.Style.Font.Size = 11; //Default font size for whole sheet
            worksheet.Cells.Style.Font.Color.SetColor(System.Drawing.Color.Black); // set color text
            worksheet.Cells.Style.Font.Name = "游ゴシック"; //Default Font name for whole sheet

            //int startIndex = 3;
            //// header clor
            try
            {
                switch(buBanType)
                {
                    case ExportExcelController.FL00R_ASSY:
                        {
                            int maxRows = 0;
                            int pageNo = Convert.ToInt32(Math.Ceiling((decimal)data.Rows.Count / 36));
                            maxRows = pageNo * 36 + pageNo * 2;  // 倍数 : 38
                            int startIndexOfPage = 0;
                            int indexData = 0;
                            // load data
                            if (sheetNo == "sheet1")
                            {
                                worksheet.PrinterSettings.PrintArea = worksheet.Cells["A:1,A:" + maxRows.ToString()];
                                for (int index = 1; index <= pageNo; index += 1)
                                {
                                    // index title
                                    startIndexOfPage = (index - 1) * 38 + 1;
                                    TitlePage(ref worksheet, buBanType, sheetNo, startIndexOfPage);
                                    // body page
                                    for (int i = 3; i <= 38; i++)
                                    {
                                        if (indexData < data.Rows.Count)
                                        {
                                            worksheet.Cells[startIndexOfPage + i - 1, 1].Value = data.Rows[indexData]["パレットNo"];
                                            worksheet.Cells[startIndexOfPage + i - 1, 2].Value = data.Rows[indexData]["ラインON"];
                                            worksheet.Cells[startIndexOfPage + i - 1, 3].Value = data.Rows[indexData]["SEQ"];
                                            worksheet.Cells[startIndexOfPage + i - 1, 4].Value = data.Rows[indexData]["部品番号"];
                                            worksheet.Cells[startIndexOfPage + i - 1, 5].Value = data.Rows[indexData]["部品略式記号"];
                                            worksheet.Cells[startIndexOfPage + i - 1, 6].Value = data.Rows[indexData][" "];
                                        }
                                        else
                                        {
                                            var value = worksheet.Cells[startIndexOfPage + i - 1, 1].Value;
                                            if ((i == 3 || i == 12 || i == 21 || i == 30) && (value == null || value.ToString() == ""))
                                            {
                                                worksheet.Cells[startIndexOfPage + i - 1, 1].Value = "パレットNo";
                                                worksheet.Cells[startIndexOfPage + i - 1, 2].Value = "ラインON";
                                                worksheet.Cells[startIndexOfPage + i - 1, 3].Value = "SEQ";
                                                worksheet.Cells[startIndexOfPage + i - 1, 4].Value = "部品番号";
                                                worksheet.Cells[startIndexOfPage + i - 1, 5].Value = "部品略式記号";
                                                worksheet.Cells[startIndexOfPage + i - 1, 6].Value = " ";
                                            }
                                            else
                                            {
                                                worksheet.Cells[startIndexOfPage + i - 1, 1].Value = "";
                                                worksheet.Cells[startIndexOfPage + i - 1, 2].Value = "";
                                                worksheet.Cells[startIndexOfPage + i - 1, 3].Value = "";
                                                worksheet.Cells[startIndexOfPage + i - 1, 4].Value = "";
                                                worksheet.Cells[startIndexOfPage + i - 1, 5].Value = "";
                                                worksheet.Cells[startIndexOfPage + i - 1, 6].Value = "";
                                            }

                                        }
                                        if (i == 3 || i == 12 || i == 21 || i == 30)
                                        {
                                            worksheet.Cells[startIndexOfPage + i - 1, 1, startIndexOfPage + i - 1, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                            worksheet.Cells[startIndexOfPage + i - 1, 1, startIndexOfPage + i - 1, 6].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                                            // merge cell 
                                            if (!worksheet.Cells[startIndexOfPage + i - 1, 5, startIndexOfPage + i - 1, 6].Merge)
                                            {
                                                worksheet.Cells[startIndexOfPage + i - 1, 5, startIndexOfPage + i - 1, 6].Merge = true;
                                            }
                                        }
                                        // border cells
                                        {
                                            worksheet.Cells[startIndexOfPage + i - 1, 1].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                            worksheet.Cells[startIndexOfPage + i - 1, 2].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                            worksheet.Cells[startIndexOfPage + i - 1, 3].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                            worksheet.Cells[startIndexOfPage + i - 1, 4].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                            worksheet.Cells[startIndexOfPage + i - 1, 5].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                            worksheet.Cells[startIndexOfPage + i - 1, 6].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                            worksheet.Cells[startIndexOfPage + i - 1, 7].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                        }
                                        indexData++;
                                    }
                                    // merge cells 
                                    if (!worksheet.Cells[startIndexOfPage + 2, 7, startIndexOfPage + 37, 7].Merge)
                                    {
                                        worksheet.Cells[startIndexOfPage + 2, 7, startIndexOfPage + 37, 7].Merge = true;
                                    }
                                    // merge パレット連番
                                    LoadMergeParetoRenban(ref worksheet, startIndexOfPage, buBanType);

                                    // Alignment is center
                                    worksheet.Cells[startIndexOfPage + 2, 1, startIndexOfPage + 37, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    worksheet.Row(index * 38).PageBreak = true;
                                }

                            }
                            else if (sheetNo == "sheet2")
                            {
                                worksheet.PrinterSettings.PrintArea = worksheet.Cells["A:1,A:" + maxRows.ToString()];
                                for (int index = 1; index <= pageNo; index += 1)
                                {
                                    // index title
                                    startIndexOfPage = (index - 1) * 38 + 1;
                                    TitlePage(ref worksheet, buBanType, sheetNo, startIndexOfPage);
                                    // body page
                                    for (int i = 3; i <= 38; i++)
                                    {
                                        if (indexData < data.Rows.Count)
                                        {
                                            worksheet.Cells[startIndexOfPage + i - 1, 1].Value = data.Rows[indexData]["パレットNo"];
                                            worksheet.Cells[startIndexOfPage + i - 1, 2].Value = data.Rows[indexData]["ラインON"];
                                            worksheet.Cells[startIndexOfPage + i - 1, 3].Value = data.Rows[indexData]["SEQ"];
                                            worksheet.Cells[startIndexOfPage + i - 1, 4].Value = data.Rows[indexData]["部品番号"];
                                            worksheet.Cells[startIndexOfPage + i - 1, 5].Value = data.Rows[indexData]["部品略式記号"];
                                            worksheet.Cells[startIndexOfPage + i - 1, 6].Value = data.Rows[indexData][" "];
                                        }
                                        else
                                        {
                                            var value = worksheet.Cells[startIndexOfPage + i - 1, 1].Value;
                                            if ((i == 3 || i == 12 || i == 21 || i == 30) && (value == null || value.ToString() == ""))
                                            {
                                                worksheet.Cells[startIndexOfPage + i - 1, 1].Value = "パレットNo";
                                                worksheet.Cells[startIndexOfPage + i - 1, 2].Value = "ラインON";
                                                worksheet.Cells[startIndexOfPage + i - 1, 3].Value = "SEQ";
                                                worksheet.Cells[startIndexOfPage + i - 1, 4].Value = "部品番号";
                                                worksheet.Cells[startIndexOfPage + i - 1, 5].Value = "部品略式記号";
                                                worksheet.Cells[startIndexOfPage + i - 1, 6].Value = " ";
                                            }
                                            else
                                            {
                                                worksheet.Cells[startIndexOfPage + i - 1, 1].Value = "";
                                                worksheet.Cells[startIndexOfPage + i - 1, 2].Value = "";
                                                worksheet.Cells[startIndexOfPage + i - 1, 3].Value = "";
                                                worksheet.Cells[startIndexOfPage + i - 1, 4].Value = "";
                                                worksheet.Cells[startIndexOfPage + i - 1, 5].Value = "";
                                                worksheet.Cells[startIndexOfPage + i - 1, 6].Value = "";
                                            }
                                        }
                                        if (i == 3 || i == 12 || i == 21 || i == 30)
                                        {
                                            var value = worksheet.Cells[startIndexOfPage + i - 1, 1].Value;

                                            worksheet.Cells[startIndexOfPage + i - 1, 1, startIndexOfPage + i - 1, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                            worksheet.Cells[startIndexOfPage + i - 1, 1, startIndexOfPage + i - 1, 6].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                                            // merge cell 
                                            if (!worksheet.Cells[startIndexOfPage + i - 1, 5, startIndexOfPage + i - 1, 6].Merge)
                                            {
                                                worksheet.Cells[startIndexOfPage + i - 1, 5, startIndexOfPage + i - 1, 6].Merge = true;
                                            }
                                        }
                                        // border cells
                                        {
                                            worksheet.Cells[startIndexOfPage + i - 1, 1].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                            worksheet.Cells[startIndexOfPage + i - 1, 2].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                            worksheet.Cells[startIndexOfPage + i - 1, 3].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                            worksheet.Cells[startIndexOfPage + i - 1, 4].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                            worksheet.Cells[startIndexOfPage + i - 1, 5].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                            worksheet.Cells[startIndexOfPage + i - 1, 6].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                            worksheet.Cells[startIndexOfPage + i - 1, 7].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                        }
                                        indexData++;
                                    }
                                    // merge cells 
                                    if (!worksheet.Cells[startIndexOfPage + 2, 7, startIndexOfPage + 37, 7].Merge)
                                    {
                                        worksheet.Cells[startIndexOfPage + 2, 7, startIndexOfPage + 37, 7].Merge = true;
                                    }
                                    // merge パレット連番
                                    LoadMergeParetoRenban(ref worksheet, startIndexOfPage, buBanType);

                                    // Alignment is center
                                    worksheet.Cells[startIndexOfPage + 2, 1, startIndexOfPage + 37, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    worksheet.Row(index * 38).PageBreak = true;
                                }
                            }
                            else
                            {
                                return false;
                            }
                            break;
                        }
                    case ExportExcelController.FRAME_ASSY:
                        {
                            int maxRows = 0;
                            int pageNo = Convert.ToInt32(Math.Ceiling((decimal)data.Rows.Count / 35));
                            maxRows = pageNo * 35 + pageNo * 2;  // 倍数 : 38
                            int startIndexOfPage = 0;
                            int indexData = 0;
                            // load data
                            if (sheetNo == "sheet1")
                            {
                                worksheet.PrinterSettings.PrintArea = worksheet.Cells["A:1,A:" + maxRows.ToString()];
                                for (int index = 1; index <= pageNo; index += 1)
                                {
                                    // index title
                                    startIndexOfPage = (index - 1) * 37 + 1;
                                    TitlePage(ref worksheet, buBanType, sheetNo, startIndexOfPage);
                                    // body page
                                    for (int i = 3; i <= 37; i++)
                                    {
                                        if (indexData < data.Rows.Count)
                                        {
                                            worksheet.Cells[startIndexOfPage + i - 1, 1].Value = data.Rows[indexData]["パレットNo"];
                                            worksheet.Cells[startIndexOfPage + i - 1, 2].Value = data.Rows[indexData]["ラインON"];
                                            worksheet.Cells[startIndexOfPage + i - 1, 3].Value = data.Rows[indexData]["SEQ"];
                                            worksheet.Cells[startIndexOfPage + i - 1, 4].Value = data.Rows[indexData]["部品番号"];
                                            worksheet.Cells[startIndexOfPage + i - 1, 5].Value = data.Rows[indexData]["部品略式記号"];
                                            worksheet.Cells[startIndexOfPage + i - 1, 6].Value = data.Rows[indexData][" "];
                                        }
                                        else
                                        {
                                            var value = worksheet.Cells[startIndexOfPage + i - 1, 1].Value;
                                            if ((i == 3 || i == 8 || i == 13 || i == 18 || i == 23 || i == 28 || i == 33) && (value == null || value.ToString() == ""))
                                            {
                                                worksheet.Cells[startIndexOfPage + i - 1, 1].Value = "パレットNo";
                                                worksheet.Cells[startIndexOfPage + i - 1, 2].Value = "ラインON";
                                                worksheet.Cells[startIndexOfPage + i - 1, 3].Value = "SEQ";
                                                worksheet.Cells[startIndexOfPage + i - 1, 4].Value = "部品番号";
                                                worksheet.Cells[startIndexOfPage + i - 1, 5].Value = "部品略式記号";
                                                worksheet.Cells[startIndexOfPage + i - 1, 6].Value = " ";
                                            }
                                            else
                                            {
                                                worksheet.Cells[startIndexOfPage + i - 1, 1].Value = "";
                                                worksheet.Cells[startIndexOfPage + i - 1, 2].Value = "";
                                                worksheet.Cells[startIndexOfPage + i - 1, 3].Value = "";
                                                worksheet.Cells[startIndexOfPage + i - 1, 4].Value = "";
                                                worksheet.Cells[startIndexOfPage + i - 1, 5].Value = "";
                                                worksheet.Cells[startIndexOfPage + i - 1, 6].Value = "";
                                            }
                                        }
                                        if (i == 3 || i == 8 || i == 13 || i == 18 || i == 23 || i == 28 || i == 33)
                                        {
                                            worksheet.Cells[startIndexOfPage + i - 1, 1, startIndexOfPage + i - 1, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                            worksheet.Cells[startIndexOfPage + i - 1, 1, startIndexOfPage + i - 1, 6].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                                            // merge cell 
                                            if (!worksheet.Cells[startIndexOfPage + i - 1, 5, startIndexOfPage + i - 1, 6].Merge)
                                            {
                                                worksheet.Cells[startIndexOfPage + i - 1, 5, startIndexOfPage + i - 1, 6].Merge = true;
                                            }
                                        }
                                        // border cells
                                        {
                                            worksheet.Cells[startIndexOfPage + i - 1, 1].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                            worksheet.Cells[startIndexOfPage + i - 1, 2].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                            worksheet.Cells[startIndexOfPage + i - 1, 3].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                            worksheet.Cells[startIndexOfPage + i - 1, 4].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                            worksheet.Cells[startIndexOfPage + i - 1, 5].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                            worksheet.Cells[startIndexOfPage + i - 1, 6].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                            worksheet.Cells[startIndexOfPage + i - 1, 7].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                        }
                                        indexData++;
                                    }
                                    // merge cells 
                                    if (!worksheet.Cells[startIndexOfPage + 2, 7, startIndexOfPage + 36, 7].Merge)
                                    {
                                        worksheet.Cells[startIndexOfPage + 2, 7, startIndexOfPage + 36, 7].Merge = true;
                                    }
                                    // merge パレット連番
                                    LoadMergeParetoRenban(ref worksheet, startIndexOfPage, buBanType);

                                    // Alignment is center
                                    worksheet.Cells[startIndexOfPage + 2, 1, startIndexOfPage + 36, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    worksheet.Row(index * 37).PageBreak = true;
                                }
                            }
                            else if (sheetNo == "sheet2")
                            {
                                worksheet.PrinterSettings.PrintArea = worksheet.Cells["A:1,A:" + maxRows.ToString()];
                                for (int index = 1; index <= pageNo; index += 1)
                                {
                                    // index title
                                    startIndexOfPage = (index - 1) * 37 + 1;
                                    TitlePage(ref worksheet, buBanType, sheetNo, startIndexOfPage);
                                    // body page
                                    for (int i = 3; i <= 37; i++)
                                    {
                                        if (indexData < data.Rows.Count)
                                        {
                                            worksheet.Cells[startIndexOfPage + i - 1, 1].Value = data.Rows[indexData]["パレットNo"];
                                            worksheet.Cells[startIndexOfPage + i - 1, 2].Value = data.Rows[indexData]["ラインON"];
                                            worksheet.Cells[startIndexOfPage + i - 1, 3].Value = data.Rows[indexData]["SEQ"];
                                            worksheet.Cells[startIndexOfPage + i - 1, 4].Value = data.Rows[indexData]["部品番号"];
                                            worksheet.Cells[startIndexOfPage + i - 1, 5].Value = data.Rows[indexData]["部品略式記号"];
                                            worksheet.Cells[startIndexOfPage + i - 1, 6].Value = data.Rows[indexData][" "];
                                        }
                                        else
                                        {
                                            var value = worksheet.Cells[startIndexOfPage + i - 1, 1].Value;
                                            if ((i == 3 || i == 8 || i == 13 || i == 18 || i == 23 || i == 28 || i == 33) && (value == null || value.ToString() == ""))
                                            {
                                                worksheet.Cells[startIndexOfPage + i - 1, 1].Value = "パレットNo";
                                                worksheet.Cells[startIndexOfPage + i - 1, 2].Value = "ラインON";
                                                worksheet.Cells[startIndexOfPage + i - 1, 3].Value = "SEQ";
                                                worksheet.Cells[startIndexOfPage + i - 1, 4].Value = "部品番号";
                                                worksheet.Cells[startIndexOfPage + i - 1, 5].Value = "部品略式記号";
                                                worksheet.Cells[startIndexOfPage + i - 1, 6].Value = " ";
                                            }
                                            else
                                            {
                                                worksheet.Cells[startIndexOfPage + i - 1, 1].Value = "";
                                                worksheet.Cells[startIndexOfPage + i - 1, 2].Value = "";
                                                worksheet.Cells[startIndexOfPage + i - 1, 3].Value = "";
                                                worksheet.Cells[startIndexOfPage + i - 1, 4].Value = "";
                                                worksheet.Cells[startIndexOfPage + i - 1, 5].Value = "";
                                                worksheet.Cells[startIndexOfPage + i - 1, 6].Value = "";
                                            }
                                        }
                                        if (i == 3 || i == 8 || i == 13 || i == 18 || i == 23 || i == 28 || i == 33)
                                        {
                                            var value = worksheet.Cells[startIndexOfPage + i - 1, 1].Value;

                                            worksheet.Cells[startIndexOfPage + i - 1, 1, startIndexOfPage + i - 1, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                            worksheet.Cells[startIndexOfPage + i - 1, 1, startIndexOfPage + i - 1, 6].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                                            // merge cell 
                                            if (!worksheet.Cells[startIndexOfPage + i - 1, 5, startIndexOfPage + i - 1, 6].Merge)
                                            {
                                                worksheet.Cells[startIndexOfPage + i - 1, 5, startIndexOfPage + i - 1, 6].Merge = true;
                                            }
                                        }
                                        // border cells
                                        {
                                            worksheet.Cells[startIndexOfPage + i - 1, 1].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                            worksheet.Cells[startIndexOfPage + i - 1, 2].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                            worksheet.Cells[startIndexOfPage + i - 1, 3].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                            worksheet.Cells[startIndexOfPage + i - 1, 4].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                            worksheet.Cells[startIndexOfPage + i - 1, 5].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                            worksheet.Cells[startIndexOfPage + i - 1, 6].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                            worksheet.Cells[startIndexOfPage + i - 1, 7].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
                                        }
                                        indexData++;
                                    }
                                    // merge cells 
                                    if (!worksheet.Cells[startIndexOfPage + 2, 7, startIndexOfPage + 36, 7].Merge)
                                    {
                                        worksheet.Cells[startIndexOfPage + 2, 7, startIndexOfPage + 36, 7].Merge = true;
                                    }
                                    // merge パレット連番
                                    LoadMergeParetoRenban(ref worksheet, startIndexOfPage, buBanType);

                                    // Alignment is center
                                    worksheet.Cells[startIndexOfPage + 2, 1, startIndexOfPage + 36, 6].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    worksheet.Row(index * 37).PageBreak = true;
                                }
                            }
                            else
                            {
                                return false;
                            }
                            break;
                        }
                    default:
                        {
                            return false;
                        }
                }
            }
            catch
            {
                throw;
            }
            
            return true;
        }
        private void TitlePage(ref ExcelWorksheet worksheet, string buBanType, string sheetNo, int index)
        {
            //worksheet.Cells.Style.Font.Size = 11; //Default font size for whole sheet
            //worksheet.Cells.Style.Font.Name = "游ゴシック"; //Default Font name for whole sheet
            
            // タイトル行が指定されているときは、タイトル行をセットする
            if (buBanType == ExportExcelController.FL00R_ASSY && sheetNo == "sheet1")
            {
                worksheet.Cells[index, 1].Value = "【出荷確認用】FLOOR ASSY";
            }
            else if (buBanType == ExportExcelController.FL00R_ASSY && sheetNo == "sheet2")
            {
                worksheet.Cells[index, 1].Value = "【生産用】FlOOR ASSY";
            }
            else if (buBanType == ExportExcelController.FRAME_ASSY && sheetNo == "sheet1")
            {
                worksheet.Cells[index, 1].Value = "【出荷確認用】FLAME ASSY";
            }
            else if (buBanType == ExportExcelController.FRAME_ASSY && sheetNo == "sheet2")
            {
                worksheet.Cells[index, 1].Value = "【生産用】FLAME  ASSY";
            }
        }

        private void LoadMergeParetoRenban(ref ExcelWorksheet worksheet, int startIndexOfPage, string buBanType)
        {
            // merge パレット連番
            // .Merge = true; //Merge columns start and end range
            //  Font should be bold
            //  Alignment is center
            //
            switch (buBanType)
            {
                case ExportExcelController.FL00R_ASSY:
                    {
                        // Block 1
                        {
                            worksheet.Cells[startIndexOfPage + 3, 1, startIndexOfPage + 10, 1].Merge = true; 
                            worksheet.Cells[startIndexOfPage + 3, 1, startIndexOfPage + 10, 1].Style.Font.Bold = true; 
                            worksheet.Cells[startIndexOfPage + 3, 1, startIndexOfPage + 10, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[startIndexOfPage + 3, 1, startIndexOfPage + 10, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            worksheet.Cells[startIndexOfPage + 3, 1, startIndexOfPage + 10, 1].Style.Font.Size = 70;
                        }
                        // Block 2
                        {
                            worksheet.Cells[startIndexOfPage + 12, 1, startIndexOfPage + 19, 1].Merge = true; 
                            worksheet.Cells[startIndexOfPage + 12, 1, startIndexOfPage + 19, 1].Style.Font.Bold = true; 
                            worksheet.Cells[startIndexOfPage + 12, 1, startIndexOfPage + 19, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[startIndexOfPage + 12, 1, startIndexOfPage + 19, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            worksheet.Cells[startIndexOfPage + 12, 1, startIndexOfPage + 19, 1].Style.Font.Size = 70;
                        }
                        // Block 3
                        {
                            worksheet.Cells[startIndexOfPage + 21, 1, startIndexOfPage + 28, 1].Merge = true; 
                            worksheet.Cells[startIndexOfPage + 21, 1, startIndexOfPage + 28, 1].Style.Font.Bold = true; 
                            worksheet.Cells[startIndexOfPage + 21, 1, startIndexOfPage + 28, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[startIndexOfPage + 21, 1, startIndexOfPage + 28, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            worksheet.Cells[startIndexOfPage + 21, 1, startIndexOfPage + 28, 1].Style.Font.Size = 70;
                        }
                        // Block 4
                        {
                            worksheet.Cells[startIndexOfPage + 30, 1, startIndexOfPage + 37, 1].Merge = true; 
                            worksheet.Cells[startIndexOfPage + 30, 1, startIndexOfPage + 37, 1].Style.Font.Bold = true; 
                            worksheet.Cells[startIndexOfPage + 30, 1, startIndexOfPage + 37, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[startIndexOfPage + 30, 1, startIndexOfPage + 37, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            worksheet.Cells[startIndexOfPage + 30, 1, startIndexOfPage + 37, 1].Style.Font.Size = 70;
                        }
                        break;
                    }
                case ExportExcelController.FRAME_ASSY:
                    {
                        // Block 1
                        {
                            worksheet.Cells[startIndexOfPage + 3, 1, startIndexOfPage + 6, 1].Merge = true;
                            worksheet.Cells[startIndexOfPage + 3, 1, startIndexOfPage + 6, 1].Style.Font.Bold = true;
                            worksheet.Cells[startIndexOfPage + 3, 1, startIndexOfPage + 6, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[startIndexOfPage + 3, 1, startIndexOfPage + 6, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            worksheet.Cells[startIndexOfPage + 3, 1, startIndexOfPage + 6, 1].Style.Font.Size = 70;
                        }
                        // Block 2
                        {
                            worksheet.Cells[startIndexOfPage + 8, 1, startIndexOfPage + 11, 1].Merge = true;
                            worksheet.Cells[startIndexOfPage + 8, 1, startIndexOfPage + 11, 1].Style.Font.Bold = true;
                            worksheet.Cells[startIndexOfPage + 8, 1, startIndexOfPage + 11, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[startIndexOfPage + 8, 1, startIndexOfPage + 11, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            worksheet.Cells[startIndexOfPage + 8, 1, startIndexOfPage + 11, 1].Style.Font.Size = 70;
                        }
                        // Block 3
                        {
                            worksheet.Cells[startIndexOfPage + 13, 1, startIndexOfPage + 16, 1].Merge = true;
                            worksheet.Cells[startIndexOfPage + 13, 1, startIndexOfPage + 16, 1].Style.Font.Bold = true;
                            worksheet.Cells[startIndexOfPage + 13, 1, startIndexOfPage + 16, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[startIndexOfPage + 13, 1, startIndexOfPage + 16, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            worksheet.Cells[startIndexOfPage + 13, 1, startIndexOfPage + 16, 1].Style.Font.Size = 70;
                        }
                        // Block 4
                        {
                            worksheet.Cells[startIndexOfPage + 18, 1, startIndexOfPage + 21, 1].Merge = true;
                            worksheet.Cells[startIndexOfPage + 18, 1, startIndexOfPage + 21, 1].Style.Font.Bold = true;
                            worksheet.Cells[startIndexOfPage + 18, 1, startIndexOfPage + 21, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[startIndexOfPage + 18, 1, startIndexOfPage + 21, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            worksheet.Cells[startIndexOfPage + 18, 1, startIndexOfPage + 21, 1].Style.Font.Size = 70;
                        }

                        // Block 5
                        {
                            worksheet.Cells[startIndexOfPage + 23, 1, startIndexOfPage + 26, 1].Merge = true;
                            worksheet.Cells[startIndexOfPage + 23, 1, startIndexOfPage + 26, 1].Style.Font.Bold = true;
                            worksheet.Cells[startIndexOfPage + 23, 1, startIndexOfPage + 26, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[startIndexOfPage + 23, 1, startIndexOfPage + 26, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            worksheet.Cells[startIndexOfPage + 23, 1, startIndexOfPage + 26, 1].Style.Font.Size = 70;
                        }
                        // Block 6
                        {
                            worksheet.Cells[startIndexOfPage + 28, 1, startIndexOfPage + 31, 1].Merge = true;
                            worksheet.Cells[startIndexOfPage + 28, 1, startIndexOfPage + 31, 1].Style.Font.Bold = true;
                            worksheet.Cells[startIndexOfPage + 28, 1, startIndexOfPage + 31, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[startIndexOfPage + 28, 1, startIndexOfPage + 31, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            worksheet.Cells[startIndexOfPage + 28, 1, startIndexOfPage + 31, 1].Style.Font.Size = 70;
                        }
                        // Block 7
                        {
                            worksheet.Cells[startIndexOfPage + 33, 1, startIndexOfPage + 36, 1].Merge = true;
                            worksheet.Cells[startIndexOfPage + 33, 1, startIndexOfPage + 36, 1].Style.Font.Bold = true;
                            worksheet.Cells[startIndexOfPage + 33, 1, startIndexOfPage + 36, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[startIndexOfPage + 33, 1, startIndexOfPage + 36, 1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            worksheet.Cells[startIndexOfPage + 33, 1, startIndexOfPage + 36, 1].Style.Font.Size = 70;
                        }

                        break;
                    }
            }

        }
        #region CSV出力
        /// <summary>
        /// CSV出力
        /// </summary>
        /// <param name="dt">出力データ</param>
        /// <param name="exportfileFullPath">出力パス(フルパス)</param>
        /// <param name="titleRows">タイトル行情報</param>
        /// <param name="quoteColumns">クォートで括る列番号</param>
        /// <returns></returns>
        public bool ExportCsv(DataTable dt, string exportfileFullPath, List<string> titleRows = null, List<int> quoteColumns = null)
        {
            // データがない時は中断
            if (dt == null || dt.Rows.Count == 0)
            {
                return false;
            }
            // 出力ファイルパスが未指定の場合は中断する
            if (String.IsNullOrWhiteSpace(exportfileFullPath))
            {
                return false;
            }
            // 出力フォルダが存在しない場合は中断する
            if (!Directory.Exists(Path.GetDirectoryName(exportfileFullPath)))
            {
                return false;
            }
            // 既にファイルが存在している場合は削除する
            if (File.Exists(exportfileFullPath))
            {
                File.Delete(exportfileFullPath);
            }
            try
            {
                var stb = new StringBuilder();
                var quoteFormat = string.Concat("\"", "{0}", "\"");

                // 出力用ファイルを生成する
                using (StreamWriter srw = new StreamWriter(exportfileFullPath))
                {
                    // タイトル行が指定されているときは、タイトル行をセットする
                    if (titleRows != null && titleRows.Count > 0)
                    {
                        for (int i = 0; i < titleRows.Count; i++)
                        {
                            // タイトル行は全列クォートで括る
                            stb.Append(",");
                            stb.Append(string.Format(quoteFormat, titleRows[i]));
                        }
                        // 書き込み
                        srw.WriteLine(stb.ToString().Substring(1));
                    }

                    for (int r = 0; r < dt.Rows.Count; r++)
                    {
                        stb.Length = 0;

                        for (int c = 0; c < dt.Columns.Count; c++)
                        {
                            stb.Append(",");
                            // クォートで括る指定のされている列は、クォートで括る
                            if (quoteColumns != null && quoteColumns.Contains(c))
                            {
                                stb.Append(string.Format(quoteFormat, NullToBlank(dt.Rows[r][c])));
                            }
                            else
                            {
                                stb.Append(NullToBlank(dt.Rows[r][c]));
                            }
                        }
                        // 書き込み
                        srw.WriteLine(stb.ToString().Substring(1));
                    }
                    srw.Flush();
                }
            }
            catch
            {
                // 失敗したときは出力用ファイルを削除する
                if (File.Exists(exportfileFullPath))
                {
                    File.Delete(exportfileFullPath);
                }

                throw;
            }

            return true;
        }
        #endregion
        #region CSV出力
        /// <summary>
        /// CSV出力
        /// </summary>
        /// <param name="dt">出力データ</param>
        /// <param name="exportfileFullPath">出力パス(フルパス)</param>
        /// <param name="titleRows">タイトル行情報</param>
        /// <param name="quoteColumns">クォートで括る列番号</param>
        /// <returns></returns>
        public bool ExportListTOCsv(List<List<string>> list, string exportfileFullPath, List<string> titleRows = null, List<int> quoteColumns = null)
        {
            // データがない時は中断
            if (list == null || list.Count == 0)
            {
                return false;
            }
            // 出力ファイルパスが未指定の場合は中断する
            if (String.IsNullOrWhiteSpace(exportfileFullPath))
            {
                return false;
            }
            // 出力フォルダが存在しない場合は中断する
            if (!Directory.Exists(Path.GetDirectoryName(exportfileFullPath)))
            {
                return false;
            }
            // 既にファイルが存在している場合は削除する
            if (File.Exists(exportfileFullPath))
            {
                File.Delete(exportfileFullPath);
            }
            try
            {

                // 出力用ファイルを生成する
                using (StreamWriter srw = new StreamWriter(exportfileFullPath))
                {
                    // タイトル行が指定されているときは、タイトル行をセットする
                    var result = string.Join(",", titleRows);
                    // 書き込み
                    srw.WriteLine(result);

                    foreach (var item in list)
                    {
                        var row = string.Join(",", item);
                        srw.WriteLine(row);
                    }
                    srw.Flush();
                }

            }
            catch
            {
                // 失敗したときは出力用ファイルを削除する
                if (File.Exists(exportfileFullPath))
                {
                    File.Delete(exportfileFullPath);
                }

                throw;
            }

            return true;
        }
        #endregion

    }
}
