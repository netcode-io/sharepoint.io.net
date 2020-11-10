using SharePoint.IO;
using System;
using System.Net;
using System.Security;

namespace Microsoft.Extensions.DependencyInjection
{
    /// <summary>
    /// SharePointServiceCollectionExtensions
    /// </summary>
    public static class SharePointServiceCollectionExtensions
    {
        class ParsedSharePointOptions : ISharePointOptions
        {
            readonly ParsedConnectionString _parsedConnectionString;
            public ParsedSharePointOptions(string connectionString) => _parsedConnectionString = new ParsedConnectionString(connectionString);
            public string Endpoint => $"https://{_parsedConnectionString.Server}";
            public NetworkCredential ServiceLogin => _parsedConnectionString.Credential;
        }

        /// <summary>
        /// Adds the SharePoint context.
        /// </summary>
        /// <param name="services">The services.</param>
        /// <param name="config">The configuration.</param>
        /// <param name="name">The name.</param>
        /// <returns></returns>
        /// <exception cref="System.ArgumentNullException">services</exception>
        public static IServiceCollection AddSharePointContext(this IServiceCollection services, ISharePointConnectionString config, string name = null)
        {
            if (services == null)
                throw new ArgumentNullException(nameof(services));
            if (config == null)
                throw new ArgumentNullException(nameof(config));
            services.Add(ServiceDescriptor.Singleton<ISharePointContext>(new SharePointContext(new ParsedSharePointOptions(config[name]))));
            return services;
        }
        /// <summary>
        /// Adds the SharePoint context.
        /// </summary>
        /// <param name="services">The services.</param>
        /// <param name="options">The configuration.</param>
        /// <returns></returns>
        /// <exception cref="System.ArgumentNullException">services</exception>
        public static IServiceCollection AddSharePointContext(this IServiceCollection services, ISharePointOptions options)
        {
            if (services == null)
                throw new ArgumentNullException(nameof(services));
            if (options == null)
                throw new ArgumentNullException(nameof(options));
            services.Add(ServiceDescriptor.Singleton<ISharePointContext>(new SharePointContext(options)));
            return services;
        }
    }
}
