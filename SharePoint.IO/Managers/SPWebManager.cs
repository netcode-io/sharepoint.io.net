using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using System;
using System.Threading.Tasks;

namespace SharePoint.IO.Managers
{
    /// <summary>
    /// SPWebManager
    /// </summary>
    public class SPWebManager
    {
        readonly ClientContext _ctx;
        readonly ILogger _log;

        /// <summary>
        /// Initializes a new instance of the <see cref="SPWebManager" /> class.
        /// </summary>
        /// <param name="ctx">The client context.</param>
        /// <param name="log">The log.</param>
        public SPWebManager(ClientContext ctx, ILogger log)
        {
            _ctx = ctx;
            _log = log;
        }

        /// <summary>
        /// Gets the current web.
        /// </summary>
        /// <value>
        /// The current web.
        /// </value>
        public Web CurrentWeb { get; private set; }

        /// <summary>
        /// Gets the current site.
        /// </summary>
        /// <value>
        /// The current site.
        /// </value>
        public Site CurrentSite { get; private set; }

        /// <summary>
        /// Loads the web asynchronous.
        /// </summary>
        /// <param name="url">The URL.</param>
        /// <returns></returns>
        public Task<Web> LoadWebAsync(string url = null) => LoadWebAsync(string.IsNullOrEmpty(url) ? _ctx.Web : _ctx.Site.OpenWeb(url));

        /// <summary>
        /// Loads the web asynchronous.
        /// </summary>
        /// <param name="web">The web.</param>
        /// <returns></returns>
        public async Task<Web> LoadWebAsync(Web web)
        {
            _ctx.Load(CurrentWeb = web ?? _ctx.Web);
            await _ctx.ExecuteQueryAsync();
            return CurrentWeb;
        }

        /// <summary>
        /// Loads the site asynchronous.
        /// </summary>
        /// <param name="site">The site.</param>
        /// <returns></returns>
        public async Task<Site> LoadSiteAsync(Site site = null)
        {
            _ctx.Load(CurrentSite = site ?? _ctx.Site);
            await _ctx.ExecuteQueryAsync();
            return CurrentSite;
        }

        /// <summary>
        /// Executes the web query asynchronous.
        /// </summary>
        /// <param name="web">The web.</param>
        public async Task ExecuteWebQueryAsync(Web web = null)
        {
            var targetWeb = web ?? CurrentWeb ?? _ctx.Web;
            targetWeb.Update();
            _ctx.Load(targetWeb);
            await _ctx.ExecuteQueryAsync();
        }

        /// <summary>
        /// Sets the web properties asynchronous.
        /// </summary>
        /// <param name="loadAllProperties">if set to <c>true</c> [load all properties].</param>
        /// <param name="action">The action.</param>
        /// <param name="loadWebAfter">if set to <c>true</c> [load web after].</param>
        /// <param name="web">The web.</param>
        public async Task SetWebPropertiesAsync(bool loadAllProperties, Action<PropertyValues, Web> action, bool loadWebAfter = false, Web web = null)
        {
            var targetWeb = web ?? CurrentWeb ?? _ctx.Web;
            if (loadAllProperties)
            {
                _ctx.Load(targetWeb.AllProperties);
                await _ctx.ExecuteQueryAsync();
            }
            action(targetWeb.AllProperties, web);
            targetWeb.Update();
            if (loadWebAfter) _ctx.Load(targetWeb);
            await _ctx.ExecuteQueryAsync();
        }
    }
}
