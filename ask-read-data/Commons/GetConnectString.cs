using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Threading.Tasks;


/// <summary>
/// 東山DBサーバに接続するための文字列を取得する。
/// </summary>
namespace ask_read_data.Commons
{
    public class GetConnectString
    {
        public string ConnectionString { get; set; }

        public GetConnectString()
        {
            var databaseName = "ask_datadb";
    #if DEBUG
            databaseName = "ask_datadb";
    #endif
            var builder = new ConfigurationBuilder()
                    .SetBasePath(Directory.GetCurrentDirectory())
                    .AddJsonFile("appsettings.json", optional: false);
            var configuration = builder.Build();
            ConnectionString = configuration.GetSection("connectionString").GetValue<string>(databaseName);
        }
    }
}
