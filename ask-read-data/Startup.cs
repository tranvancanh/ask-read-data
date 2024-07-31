using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.CookiePolicy;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.HttpsPolicy;
using Microsoft.AspNetCore.Mvc.Authorization;
using Microsoft.AspNetCore.Routing;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using ask_read_data.Areas.Admin.Models;
using ask_read_data.Areas.Admin.Repository;
using ask_read_data.Areas.Admin.Servive;
using read_data.Areas.Admin.Servive;
using ask_read_data.Repository;
using ask_read_data.Servive;
using ask_read_data.Models;
using ask_read_data.Models.Entity;
using ask_read_data.Commons;

namespace ask_read_data
{
    public class Startup
    {
        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;
        }

        public IConfiguration Configuration { get; }

        // This method gets called by the runtime. Use this method to add services to the container.
        public void ConfigureServices(IServiceCollection services)
        {
            services.AddControllers();

            //services.AddControllersWithViews();
#if DEBUG
            services.AddControllersWithViews().AddRazorRuntimeCompilation();
#else
            services.AddControllersWithViews();
#endif

            // セッションを使う
            //services.AddSession(options => {
            //    // セッションクッキーの名前を変えるなら
            //    options.Cookie.Name = "session";
            //});
            services.Configure<CookiePolicyOptions>(options =>
            {
                options.CheckConsentNeeded = context => !context.User.Identity.IsAuthenticated;
                options.MinimumSameSitePolicy = SameSiteMode.None;
                options.Secure = CookieSecurePolicy.Always;
                options.HttpOnly = HttpOnlyPolicy.Always;
            });
            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // クッキー認証に必要なサービスを登録
            services.AddAuthentication(options =>
            {
                options.DefaultAuthenticateScheme = CookieAuthenticationDefaults.AuthenticationScheme;
                options.DefaultScheme = CookieAuthenticationDefaults.AuthenticationScheme;
                options.DefaultSignInScheme = CookieAuthenticationDefaults.AuthenticationScheme;
                options.DefaultSignOutScheme = CookieAuthenticationDefaults.AuthenticationScheme;
                options.DefaultForbidScheme = CookieAuthenticationDefaults.AuthenticationScheme;
            })
            .AddCookie(CookieAuthenticationDefaults.AuthenticationScheme, options =>
            {

                // リダイレクトするログインURLも小文字に変える
                options.LoginPath = CookieAuthenticationDefaults.LoginPath.ToString().ToLower();
                options.Cookie.IsEssential = true;
                options.Cookie.HttpOnly = true;
                options.Cookie.SecurePolicy = CookieSecurePolicy.Always;
                options.Cookie.SameSite = SameSiteMode.None;
                options.Cookie.Name = CookieAuthenticationDefaults.AuthenticationScheme;
                options.Cookie.MaxAge = TimeSpan.FromMinutes(1440);
                options.LoginPath = new PathString("/Login/Login/");
                options.SlidingExpiration = true;
                options.ExpireTimeSpan = TimeSpan.FromMinutes(1440);

            });
            services.AddAuthorization(options =>
            {
                // AllowAnonymous 属性が指定されていないすべての Action などに対してユーザー認証が必要となる
                options.FallbackPolicy = new AuthorizationPolicyBuilder()
                  .RequireAuthenticatedUser()
                  .Build();
            });
            //services.Configure<RouteOptions>(options =>
            //{
            //    // URLは小文字にする
            //    options.LowercaseUrls = true;
            //});
            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            services.AddHttpContextAccessor();
            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // MVCで利用するサービスを登録
            services.AddMvc(options =>
            {
                // グローバルフィルタに承認フィルタを追加
                // すべてのコントローラでログインが必要にしておく
                var policy = new AuthorizationPolicyBuilder()
                    .RequireAuthenticatedUser()
                    .Build();
                options.Filters.Add(new AuthorizeFilter(policy));
                options.EnableEndpointRouting = false;
                
            });
            services.AddSession();
            //services.AddSingleton<IHttpContextAccessor, HttpContextAccessor>();
            services.AddTransient<IHttpContextAccessor, HttpContextAccessor>();
            services.Add(new ServiceDescriptor(typeof(ILoginViewModel), typeof(LoginService), ServiceLifetime.Transient));
            services.Add(new ServiceDescriptor(typeof(IUser), typeof(UserService), ServiceLifetime.Transient));
            services.Add(new ServiceDescriptor(typeof(IRegisterService), typeof(RegisterService), ServiceLifetime.Transient)); 
            services.Add(new ServiceDescriptor(typeof(UserViewModel), typeof(UserViewModel), ServiceLifetime.Transient));
            services.Add(new ServiceDescriptor(typeof(IImportData), typeof(ImportDataService), ServiceLifetime.Transient));
            services.Add(new ServiceDescriptor(typeof(DataModel), typeof(DataModel), ServiceLifetime.Transient)); 
            services.Add(new ServiceDescriptor(typeof(Bu_MastarModel), typeof(Bu_MastarModel), ServiceLifetime.Transient)); 
            services.Add(new ServiceDescriptor(typeof(IData), typeof(DataService), ServiceLifetime.Transient));
            services.Add(new ServiceDescriptor(typeof(IExportExcel), typeof(ExportExcelService), ServiceLifetime.Transient));
            services.AddScoped<GetConnectString>();
            services.AddHostedService<StartupTaskService>();

            services.AddRazorPages();
            services.AddAuthorization(options =>
            {
                options.AddPolicy("1", policy =>
                {
                    policy.RequireClaim(ClaimTypes.Role, "1");
                });
                options.AddPolicy("test", policy =>
                {
                    policy.RequireClaim(ClaimTypes.Role, "test");
                });
            });

            services.AddHttpContextAccessor(); // Required for IHttpContextAccessor
        }

        // This method gets called by the runtime. Use this method to configure the HTTP request pipeline.
        public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
        {
            if (env.IsDevelopment())
            {
                app.UseDeveloperExceptionPage();
            }
            else
            {
                app.UseExceptionHandler("/Home/Error");
                // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
                app.UseHsts();
            }
            app.UseHttpsRedirection();
            app.UseStaticFiles();


            // Cookieの原則機能を有効にする
            app.UseCookiePolicy();
            // IDを有効にする
            app.UseAuthentication();

            app.UseRouting();

            //認証機能を有効にします
            app.UseAuthorization();
            //これらの3つの前後の順序を逆にすることはできません
            app.UseSession();
            app.UseMvc();
            app.UseEndpoints(endpoints =>
            {
                //endpoints.MapControllers();
                //endpoints.MapRazorPages();
                endpoints.MapAreaControllerRoute(
                    name: "Admin",
                    areaName: "Admin",
                    pattern: "{controller=Login}/{action=Login}/{id?}");

                endpoints.MapControllerRoute(
                    name: "default",
                    pattern: "{controller=Home}/{action=Index}/{id?}");
            });
        }
    }
}
