using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration;
using System;
using System.IO;


/// <summary>
/// 東山DBサーバに接続するための文字列を取得する。
/// </summary>
namespace ask_read_data.Commons
{
    public class GetConnectString
    {

        public string ConnectionString()
        {
            try
            {
                var _httpContextAccessor = new HttpContextAccessor();
                var request = _httpContextAccessor.HttpContext.Request;
                var host = request.Host.Value;
                var databaseName = string.Empty;
                string pathBase = _httpContextAccessor.HttpContext.Request.PathBase;
                string path = _httpContextAccessor.HttpContext.Request.Path;
                WriteLog("pathBase :" + pathBase);
                WriteLog("path :" + path);
                WriteLog("host :" + host);
#if DEBUG
                databaseName = "ask_datadb_20240729111819";
#else
                if (host.Equals("www.tozan.co.jp", StringComparison.OrdinalIgnoreCase))
                {
                    if (pathBase.Equals("/ask-read-data"))
                    {
                        databaseName = "ask_datadb"; // ConnectionString for 本番環境
                    }
                    else if (pathBase.Equals("/ask-read-data-test"))
                    {
                        databaseName = "ask_datadb_20240729111819"; // ConnectionString for the test テスト環境
                    }
                    else 
                    {
                        databaseName = string.Empty;
                    }
                }
                else
                {
                    databaseName = "ask_datadb_20240729111819";
                }
#endif

                var builder = new ConfigurationBuilder()
                       .SetBasePath(Directory.GetCurrentDirectory())
                       .AddJsonFile("appsettings.json", optional: false);
                var configuration = builder.Build();
                var connectionString = configuration.GetSection("connectionString").GetValue<string>(databaseName); // Default case
                WriteLog("connectionString:" + connectionString);
                return connectionString;
            }
            catch (Exception ex)
            {
                WriteLog("Exception:" + ex.Message);
                return "";
            }

        }

        public void WriteLog(string strLog)
        {
            var rootPath = Directory.GetCurrentDirectory();
            var logFilePath = Path.Combine(rootPath, @"wwwroot\Logs");
            if (!System.IO.Directory.Exists(logFilePath))
            {
                System.IO.Directory.CreateDirectory(logFilePath);
            }

            var filePath = Path.Combine(logFilePath, "Log-" + System.DateTime.Today.ToString("yyyy-MM-dd") + "." + "log");
            if (!System.IO.File.Exists(filePath))
                using (System.IO.File.CreateText(filePath)) { }

            var striDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.ffff");
            File.AppendAllText(filePath, striDate + " | " + strLog + Environment.NewLine);

            //FileInfo logFileInfo = new FileInfo(filePath);
            //DirectoryInfo logDirInfo = new DirectoryInfo(logFileInfo.DirectoryName);
            //if (!logDirInfo.Exists) logDirInfo.Create();

            //using (FileStream fileStream = new FileStream(filePath, FileMode.Append))
            //{
            //    using (StreamWriter log = new StreamWriter(fileStream))
            //    {
            //        log.WriteLine(strLog);
            //    }
            //}
        }

    }
}
