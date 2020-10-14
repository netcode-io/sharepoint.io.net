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
            public ParsedSharePointOptions(ISharePointConnectionString config) => _parsedConnectionString = new ParsedConnectionString(config.String);
            public string Endpoint => $"https://{_parsedConnectionString.Server}";
            public NetworkCredential ServiceLogin => _parsedConnectionString.Credential;
        }

        /// <summary>
        /// Adds the SharePoint context.
        /// </summary>
        /// <param name="services">The services.</param>
        /// <param name="config">The configuration.</param>
        /// <returns></returns>
        /// <exception cref="System.ArgumentNullException">services</exception>
        public static IServiceCollection AddO365Context(this IServiceCollection services, ISharePointConnectionString config)
        {
            if (services == null)
                throw new ArgumentNullException(nameof(services));
            services.Add(ServiceDescriptor.Singleton<ISharePointContext>(new SharePointContext(new ParsedSharePointOptions(config))));
            return services;
        }
        /// <summary>
        /// Adds the SharePoint context.
        /// </summary>
        /// <param name="services">The services.</param>
        /// <param name="config">The configuration.</param>
        /// <returns></returns>
        /// <exception cref="System.ArgumentNullException">services</exception>
        public static IServiceCollection AddSharePointContext(this IServiceCollection services, ISharePointOptions config)
        {
            if (services == null)
                throw new ArgumentNullException(nameof(services));
            services.Add(ServiceDescriptor.Singleton<ISharePointContext>(new SharePointContext(config)));
            return services;
        }
    }
}
