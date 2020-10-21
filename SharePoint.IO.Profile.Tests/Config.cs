using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using System;

namespace SharePoint.IO.Profile
{
    internal class Config
    {
        public static IConfiguration Configuration { get; } = new ConfigurationBuilder().AddJsonFile("appsettings.json", true, true).Build();
        public static IServiceProvider Services { get; } = ConfigureServices(new ServiceCollection()).BuildServiceProvider();
        public static IServiceCollection ConfigureServices(IServiceCollection services) => services;
    }
}
