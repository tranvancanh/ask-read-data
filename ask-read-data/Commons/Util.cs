using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

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
}
