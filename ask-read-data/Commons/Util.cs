using System;
using System.Data;
using System.Data.SqlClient;
using System.Text;

namespace ask_tzn_funamiKD.Commons
{
    public class Util
    {
        public static string NullToBlank(string value)
        {
            // NULL、DBNullのときは空文字に変換する
            if (string.IsNullOrWhiteSpace(value) == true)
            {
                return string.Empty;
            }
            else
            {
                return Convert.ToString(value).Trim();
            }
            
        }
        public static int NullToBlank(object value)
        {
            // NULL、DBNullのときは空文字に変換する
            if (string.IsNullOrWhiteSpace(value.ToString()) == true)
            {
                return 0;
            }
            else
            {
                return Convert.ToInt32(value.ToString().Trim());
            }
        }
        public static decimal Decimal(object value)
        {
            // NULL、DBNullのときは空文字に変換する
            if (string.IsNullOrWhiteSpace(value.ToString()) == true)
            {
                return (decimal)0;
            }
            else
            {
                return Convert.ToDecimal(value);
            }
        }
        public static double DivideTwoNumbers(int a, int b)
        {
            // ゼロのときは空文字に変換する
            if (b == 0)
            {
                return 0;
            }
            else
            {
                return Math.Ceiling((double)a/b);
            }

        }
        public static string GetDataDB(SqlDataReader reader, string column)
        {
            if (string.IsNullOrWhiteSpace(reader[column].ToString()) == true)
            {
                return reader[column].ToString().Trim();
            }
            else
            {
                return string.Empty;
            }
        }
    }

    public static class DataTableExtensions
    {
        public static string ToCsv(this DataTable dataTable, char delimiter = ',')
        {
            StringBuilder csvData = new StringBuilder();

            // Write the header line
            foreach (DataColumn column in dataTable.Columns)
            {
                csvData.Append(column.ColumnName + delimiter);
            }
            csvData.Length--; // Remove the last delimiter
            csvData.Length--; // Remove the last delimiter
            csvData.Length--; // Remove the last delimiter
            csvData.AppendLine();

            // Write data rows
            foreach (DataRow row in dataTable.Rows)
            {
                foreach (DataColumn column in dataTable.Columns)
                {
                    csvData.Append(row[column].ToString() + delimiter);
                }
                csvData.Length--; // Remove the last delimiter
                csvData.Length--; // Remove the last delimiter
                csvData.AppendLine();
            }

            return csvData.ToString();
        }
    }
}
