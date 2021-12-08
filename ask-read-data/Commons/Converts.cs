using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ask_tzn_funamiKD.Commons
{
    public static class Converts
    {
        public static HashSet<string> ConvertToHashSet(string str)
        {
            HashSet<string> HashSetStr = new HashSet<string>();
            var value = Util.NullToBlank(str);
            if (value != string.Empty)
                HashSetStr.Add(value);

            return HashSetStr;
        }
        public static DateTimeDetail ConverDateTime(DateTime dateTime)
        {
            DateTimeDetail Details;
            int month = dateTime.Month;
            if (!((month >= 1) && (month <= 12)))
                throw new Exception("範囲外の時間");
            int year = dateTime.Year;
            if (!((year >= 1900) && (year <= 9999)))
                throw new Exception("範囲外の時間");
            int day = dateTime.Day;
            if (!((day >= 1) && (day <= 31)))
                throw new Exception("範囲外の時間");
            string day1 = (day < 10) ? ('0' + day.ToString()) : day.ToString();
            string month1 = (month < 10) ? ('0' + month.ToString()) : month.ToString();
            string jun1 = "1";
            string monthYear = year.ToString() + month1;
            if (day >= 1 && day <= 10)
                jun1 = "1";
            else if (day >= 11 && day <= 20)
                jun1 = "2";
            else if (day >= 21 && day <= 31)
                jun1 = "3";
            else
                throw new Exception("範囲外の時間");
            Details = new DateTimeDetail(year.ToString(), monthYear, jun1, month1, day1);
            return Details;
        }
        /// <summary>
        /// Max 1000Jun
        /// </summary>
        /// <param name="Datime"></param>
        /// <param name="Jun"></param>
        /// <returns></returns>
        public static DateTime AddJuns(DateTime Datime, int Jun)
        {
            DateTime newDateTime = new DateTime();
            if (Jun == 0)
                return Datime;
            if(Jun > 0)
            {

            }

            return newDateTime;
        }
        public static int NextMonth(DateTime dateTime)
        {
            int ThisMonth = dateTime.Month;
            int NextMonth = dateTime.AddMonths(1).Month;
            return NextMonth;
        }
        public static int RoundUp(in int a, in int b)
        {
            double result = 0;
            if (b == 0)
                throw new Exception("分母がゼロだった");
            if (a < 0)
                throw new Exception("範囲外の数");
            if (b < 0)
                throw new Exception("範囲外の数");
            result = (double)a / b;

            return Convert.ToInt32(Math.Ceiling(result));
        }
        public static int RoundDown(in int a, in int b)
        {
            double result = 0.00;
            if (b == 0)
                throw new Exception("分母がゼロだった");
            if (a < 0)
                throw new Exception("範囲外の数");
            if (b < 0)
                throw new Exception("範囲外の数");
            result = (double)a / b;

            return Convert.ToInt32(Math.Floor(result));
        }
        // return Last Day In Month
        public static int GetLastDayInMonth(int year, int month)
        {
            DateTime aDateTime = new DateTime(year, month, 1);

            // add 1 month and - 1 day
            int aDay = aDateTime.AddMonths(1).AddDays(-1).Day;
            return aDay;
        }
        // 1日の開始時刻を取得する
        public static DateTime StartOfDay(this DateTime theDate)
        {
            // return theDate.Date;
            return new DateTime(theDate.Year, theDate.Month, theDate.Day, 00, 00, 00, 000);
        }
        // 1日の終了時刻を取得する
        public static DateTime EndOfDay(this DateTime theDate)
        {
            return theDate.Date.AddDays(1).AddTicks(-1);
        }
        public static DateTime StartOfTheDay(this DateTime d) => new DateTime(d.Year, d.Month, d.Day, 00, 00, 00, 000);
        public static DateTime EndOfTheDay(this DateTime d) => new DateTime(d.Year, d.Month, d.Day, 23, 59, 59, 999);

        //public static bool DeepEquals(this OutPutHand obj, OutPutHand another)
        //{
        //    //Nếu null hoặc giống nhau, trả true
        //    if (ReferenceEquals(obj, another)) return true;

        //    //Nếu 1 trong 2 là null, trả false
        //    if ((obj == null) || (another == null)) return false;

        //    return obj.OutputHand.Trcd.Equals(another.OutputHand.Trcd)
        //           && obj.OutputHand.Hinban.Equals(another.OutputHand.Hinban)
        //           && obj.OutputHand.Kotei.Equals(another.OutputHand.Kotei)
        //           && obj.OutputHand.HandValue == another.OutputHand.HandValue
        //           && obj.OutputHand.HandUpDateTime.Equals(another.OutputHand.HandUpDateTime);
        //}
        public static bool JSONEquals(this object obj, object another)
        {
            if (ReferenceEquals(obj, another)) return true;
            if ((obj == null) || (another == null)) return false;
            if (obj.GetType() != another.GetType()) return false;

            var objJson = JsonConvert.SerializeObject(obj);
            var anotherJson = JsonConvert.SerializeObject(another);

            return objJson == anotherJson;
        }
    }
    public class DateTimeDetail
    {
        public string Year { get; set; }
        public string MonthYear { get; set; }
        public string Jun { get; set; }
        public string Month { get; set; }
        public string Day { get; set; }

        public DateTimeDetail(string Year, string MonthYear, string Jun, string Month, string Day)
        {
            this.Year = Year;
            this.MonthYear = MonthYear;
            this.Jun = Jun;
            this.Month = Month;
            this.Day = Day;
        }
    }
}
