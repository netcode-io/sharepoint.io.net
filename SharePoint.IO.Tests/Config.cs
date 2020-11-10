using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using System;

namespace SharePoint.IO
{
    internal class Config : ISharePointConnectionString
    {
        public static IConfiguration Configuration { get; } = new ConfigurationBuilder().AddJsonFile("appsettings.json", true, true).Build();
        public static IServiceProvider Services { get; } = ConfigureServices(new ServiceCollection()).BuildServiceProvider();
        public static IServiceCollection ConfigureServices(IServiceCollection services) => services.AddSharePointContext(new Config());
        string ISharePointConnectionString.this[string name] => Configuration.GetConnectionString("SharePoint");
    }
}
