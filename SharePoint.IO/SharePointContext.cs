using Microsoft.SharePoint.Client;
using System.Net;

namespace SharePoint.IO
{
    /// <summary>
    /// ISharePointContext
    /// </summary>
    public interface ISharePointContext
    {
        /// <summary>
        /// Connects the specified subsite.
        /// </summary>
        /// <param name="subsite">The subsite.</param>
        /// <returns></returns>
        ClientContext Connect(string subsite = null);
    }

    /// <summary>
    /// SharePointContext
    /// </summary>
    public class SharePointContext : ISharePointContext
    {
        readonly string _endpoint;
        readonly NetworkCredential _credential;

        /// <summary>
        /// Initializes a new instance of the <see cref="SharePointContext"/> class.
        /// </summary>
        /// <param name="config">The configuration.</param>
        public SharePointContext(ISharePointOptions config)
            : this(config.Endpoint, config.ServiceLogin) { }
        SharePointContext(string endpoint, NetworkCredential credential)
        {
            _endpoint = endpoint;
            _credential = credential;
        }

        /// <summary>
        /// Connects the specified subsite.
        /// </summary>
        /// <param name="subsite">The subsite.</param>
        /// <returns>ClientContext.</returns>
        public ClientContext Connect(string subsite = null) => new ClientContext($"{_endpoint}{subsite}")
        {
            Credentials = new SharePointOnlineCredentials(_credential.UserName, _credential.Password)
        };
    }
}
