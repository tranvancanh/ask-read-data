using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;

namespace ask_read_data.Commons
{
    public class ErrorInfor
    {
        public static void DebugWriteLineError(Exception ex)
        {
            System.Console.ForegroundColor = ConsoleColor.Red;
            Debug.WriteLine("=========================== ここからのエラー情報 ============================");
            Debug.WriteLine("=========================== ここからのエラー情報 ============================");
            Debug.WriteLine(ex.Message);
            Debug.WriteLine(ex.StackTrace);
            Debug.WriteLine(ex.GetType().ToString());
            Debug.WriteLine("=========================== 終わり ============================");
            System.Console.ResetColor();
        }
    }
}
