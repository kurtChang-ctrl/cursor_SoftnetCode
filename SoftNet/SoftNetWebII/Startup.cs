using Base;
using Base.Enums;
using Base.Models;
using Base.Services;
using BaseWeb.Services;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc.Razor;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.FileProviders;
using Microsoft.Extensions.Hosting;
using SoftNetWebII.Services;
using System;
using System.Data.Common;
using Microsoft.Data.SqlClient;
using System.IO;
using System.Threading.Tasks;

namespace SoftNetWebII
{
    public class Startup
    {
        public Startup(IConfiguration configuration)
        {
            Configuration = configuration;//Configuration會包含appsettings.json資料, 如:Configuration["AllowedHosts"] 取AllowedHosts值
        }
        //先執行 ConfigureServices, 在Configure
        public IConfiguration Configuration { get; }

        // 建置Web所需的服務
        public void ConfigureServices(IServiceCollection services)
        {

            //1.config MVC
            services.AddControllersWithViews()
                //view Localization
                .AddViewLocalization(LanguageViewLocationExpanderFormat.Suffix)
                //use pascal for newtonSoft json
                .AddNewtonsoftJson(opts => { opts.UseMemberCasing(); })
                //use pascal for MVC json
                .AddJsonOptions(opts => { opts.JsonSerializerOptions.PropertyNamingPolicy = null; });

            //2.set Resources path
            services.AddLocalization(opts => opts.ResourcesPath = "Resources");

            //3.http context
            services.AddSingleton<IHttpContextAccessor, HttpContextAccessor>();

            //4.user info for base component
            services.AddSingleton<IBaseUserService, MyBaseUserService>();

            //4.user info for base component
            //services.AddSingleton<IBaseUserService, BaseUserService>();

            //5.ado.net for mssql
            services.AddTransient<DbConnection, SqlConnection>();
            services.AddTransient<DbCommand, SqlCommand>();

            //6.appSettings "FunConfig" section -> _Fun.Config
            var config = new ConfigDto();
            Configuration.GetSection("FunConfig").Bind(config);
            // AIReportDDay：支援單一 "Sunday" 或多天 "Monday,Wednesday"（逗號分隔），明確從字串解析以正確綁定到 DaysOfWeek
            var aiReportDDayStr = Configuration.GetSection("FunConfig")["AIReportDDay"]?.Trim();
            if (!string.IsNullOrEmpty(aiReportDDayStr))
            {
                DaysOfWeek combined = DaysOfWeek.None;
                foreach (var part in aiReportDDayStr.Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    if (Enum.TryParse<DaysOfWeek>(part.Trim(), true, out var day))
                        combined |= day;
                }
                config.AIReportDDay = combined;
            }
            _Fun.Config = config;
            _Fun.RUNProcessSYSName = System.Diagnostics.Process.GetCurrentProcess().ProcessName;


            services.AddHostedService<RUNTimeServer>();


            //7.新增services
            //20240314取消  services.AddSingleton(typeof(SocketClientService), new SocketClientService()); // 建立Socket Client 5431主動與RMS Service連線  交換工站狀態
            services.AddSingleton(typeof(SNWebSocketService), new SNWebSocketService()); // 建立WebSocket Server 開通 _Fun.Config.WesocketPort port
            services.AddSingleton(typeof(SFC_Common), new SFC_Common("1", _Fun.Config.Db)); // 建立WebSocket Server 開通 _Fun.Config.WesocketPort port
                                                                                            //services.AddSingleton(typeof(RUNTimeServer), new RUNTimeServer()); // 建立定時,定期工作 Service , 目前定60秒輪巡, 雷同TMService

            // 註冊 RagVectorStoreService 為 Singleton，並在啟動時初始化
            var ragVectorStoreService = new RagVectorStoreService();
            services.AddSingleton(typeof(RagVectorStoreService), ragVectorStoreService);


            //8.session (memory cache)
            services.AddDistributedMemoryCache();
            services.AddSession(opts =>
            {
                opts.Cookie.HttpOnly = true;
                opts.Cookie.IsEssential = true;
                opts.IdleTimeout = TimeSpan.FromHours(10);
                //opts.IdleTimeout = TimeSpan.FromSeconds(30);
            });

        }

        // 設定http請求的處置設定
        public void Configure(IApplicationBuilder app, IWebHostEnvironment env)
        {
            //1.initial & set locale  只會被run一次
            _Fun.Init(env.IsDevelopment(), app.ApplicationServices, DbTypeEnum.MSSql, AuthTypeEnum.Data);

            //2.set locale, call async method here !!
            _Locale.SetCulture(_Fun.Config.Locale);

            //3. SQL Server 向量功能：依設定 UseSQLVector 決定是否測試並啟用
            try
            {
                if (!_Fun.Config.UseSQLVector)
                {
                    _Fun.SupportsSQLVector = false;
                    _Log.InfoAsync("系統啟動：UseSQLVector=false，不使用 SQL Server 原生 vector").Wait();
                }
                else
                {
                    var ragVectorStoreService = app.ApplicationServices.GetRequiredService<RagVectorStoreService>();
                    var testTask = ragVectorStoreService.TestSQLVectorSupportAsync();
                    if (testTask.Wait(TimeSpan.FromSeconds(30)))
                    {
                        var supportsSQLVector = testTask.Result;
                        _Log.InfoAsync($"系統啟動：SQL Server 向量功能支援狀態 = {supportsSQLVector}").Wait();
                    }
                    else
                    {
                        _Fun.SupportsSQLVector = false;
                        _Log.ErrorAsync("系統啟動：SQL Server 向量功能測試逾 30 秒未完成，設為 false", false).Wait();
                    }
                }
            }
            catch (Exception ex)
            {
                _Fun.SupportsSQLVector = false;
                _Log.ErrorAsync($"系統啟動：SQL Server 向量功能錯誤: {ex.Message}，設為 false", false).Wait();
            }

            //4.exception handle
            if (env.IsDevelopment())
            {
                DeveloperExceptionPageOptions devEx = new DeveloperExceptionPageOptions
                {
                    SourceCodeLineCount = 10
                };
                app.UseDeveloperExceptionPage(devEx);
                //app.UseDeveloperExceptionPage();
                //app.UseExceptionHandler("/Home/Error"); //temp
            }
            else
            {
                app.UseExceptionHandler("/Home/Error");
                // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.
                app.UseHsts();
            }
            // 在生產環境啟用 HTTPS 重定向
            if (!env.IsDevelopment())
            {
                app.UseHttpsRedirection();
            }
            app.UseStaticFiles();//設定使用靜態檔案
            app.UseFileServer();//設定使用靜態檔案

            app.UseRouting();
            //app.UseAuthorization();

            app.UseSession();
            app.UseEndpoints(endpoints =>
            {
                endpoints.MapControllerRoute(
                    name: "default",
                    pattern: "{controller=Home}/{action=Index}/{id?}");
            });

        }
    }
}
