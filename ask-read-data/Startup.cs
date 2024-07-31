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

            // �Z�b�V�������g��
            //services.AddSession(options => {
            //    // �Z�b�V�����N�b�L�[�̖��O��ς���Ȃ�
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
            // �N�b�L�[�F�؂ɕK�v�ȃT�[�r�X��o�^
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

                // ���_�C���N�g���郍�O�C��URL���������ɕς���
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
                // AllowAnonymous �������w�肳��Ă��Ȃ����ׂĂ� Action �Ȃǂɑ΂��ă��[�U�[�F�؂��K�v�ƂȂ�
                options.FallbackPolicy = new AuthorizationPolicyBuilder()
                  .RequireAuthenticatedUser()
                  .Build();
            });
            //services.Configure<RouteOptions>(options =>
            //{
            //    // URL�͏������ɂ���
            //    options.LowercaseUrls = true;
            //});
            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            services.AddHttpContextAccessor();
            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            // MVC�ŗ��p����T�[�r�X��o�^
            services.AddMvc(options =>
            {
                // �O���[�o���t�B���^�ɏ��F�t�B���^��ǉ�
                // ���ׂẴR���g���[���Ń��O�C�����K�v�ɂ��Ă���
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


            // Cookie�̌����@�\��L���ɂ���
            app.UseCookiePolicy();
            // ID��L���ɂ���
            app.UseAuthentication();

            app.UseRouting();

            //�F�؋@�\��L���ɂ��܂�
            app.UseAuthorization();
            //������3�̑O��̏������t�ɂ��邱�Ƃ͂ł��܂���
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
