using Base.Enums;
using Base.Models;
using Base.Services;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Data.Common;
using System.Data.SqlClient;

namespace EasyRpt
{
    class Program
    {
        static void Main(string[] args)
        {
            IConfiguration configBuild = new ConfigurationBuilder()
                .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
                .Build();

            //appSettings "FunConfig" section -> _Fun.Config
            var config = new ConfigDto();
            configBuild.GetSection("FunConfig").Bind(config);
            _Fun.Config = config;

            //setup our DI
            var services = new ServiceCollection();
                //.BuildServiceProvider();

            //locale & user info for base component
            services.AddSingleton<IBaseResService, BaseResService>();
            services.AddSingleton<IBaseUserService, BaseUserService>();

            //ado.net for mssql
            services.AddTransient<DbConnection, SqlConnection>();
            services.AddTransient<DbCommand, SqlCommand>();

            //initial _Fun by mssql
            IServiceProvider di = services.BuildServiceProvider();
            _Fun.Init(di, DbTypeEnum.MSSql);

            //Console.WriteLine("Hello World!");
            _Log.Info("EasyRpt Start.");
            new MyService().Run();
            _Log.Info("EasyRpt End.");
        }
    }
}
