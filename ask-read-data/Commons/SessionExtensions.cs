using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Newtonsoft.Json;

namespace ask_read_data.Commons
{

        //public class Startup
        //{
        //    public void ConfigureServices(IServiceCollection services)
        //    {
        //        // いらないのかも
        //        //services.AddDistributedMemoryCache();

        //        // セッションを使う
        //        services.AddSession(options => {
        //            // セッションクッキーの名前を変えるなら
        //            options.Cookie.Name = "session";
        //        });

        //        services.AddMvc();
        //    }

        //    public void Configure(IApplicationBuilder app, IHostingEnvironment env)
        //    {
        //        // パイプラインでセッションを使う
        //        app.UseSession();

        //        app.UseMvcWithDefaultRoute();
        //    }
        //}
        ////////////////////////////////////////////////////////////////////////////////////////////
        // 文字列の場合
        // SetStringとGetStringの拡張メソッドを使うのに必要
      

        //// セッションに文字列を書き込む（キーは"key"）
        //HttpContext.Session.SetString("key", "Hoge");

        //// セッションから文字列を読み込む
        //HttpContext.Session.GetString("key"); 
        /// ////////////////////////////////////////////////////////////////////////////////////////
        /// <summary>
        /// //////////////////////////////////////////////////////////////////////////
        /// </summary>
        // セッションにオブジェクトを設定・取得する拡張メソッドを用意する
        public static class SessionExtensions
        {
        // セッションにオブジェクトを書き込む
        public static void SetObject<TObject>(this ISession session, string key, TObject obj)
        {
            var json = JsonConvert.SerializeObject(obj);
            session.SetString(key, json);
        }

        // セッションからオブジェクトを読み込む
        public static TObject GetObject<TObject>(this ISession session, string key)
        {
            var json = session.GetString(key);
            return string.IsNullOrEmpty(json)
                ? default(TObject)
                : JsonConvert.DeserializeObject<TObject>(json);
        }

        /// <summary>
        /// //////////////  SEt Session ////////////////////////
        /// </summary>
        // 拡張メソッドを使ってみる
        // セッションにSampleクラスを書き込む
        //HttpContext.Session.SetObject("key", new SessionExtensions());

        //// セッションからSampleクラスを読み込む
        //HttpContext.Session.GetObject<Sample>("key");
    
}
}
