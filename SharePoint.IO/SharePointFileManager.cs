using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using SharePoint.IO.Managers;
using System;
using System.Threading.Tasks;

namespace SharePoint.IO
{
    /// <summary>
    /// SharePointFileManager
    /// </summary>
    public class SharePointFileManager
    {
        readonly ClientContext _context;
        readonly Web _web;
        readonly ILogger _log;

        /// <summary>
        /// Initializes a new instance of the <see cref="SharePointFileManager" /> class.
        /// </summary>
        /// <param name="context">The context.</param>
        /// <param name="url">The URL.</param>
        /// <param name="log">The log.</param>
        public SharePointFileManager(ClientContext context, string url = null, ILogger log = null)
        {
            _context = context;
            _web = new SPWebManager(_context, log).LoadWebAsync(url).ConfigureAwait(false).GetAwaiter().GetResult();
            _log = log;
            Files = new FileShaman(_web, log);
            Folders = new FolderShaman(_web, log);
        }

        /// <summary>
        /// Gets the files.
        /// </summary>
        /// <value>
        /// The files.
        /// </value>
        public FileShaman Files { get; }

        /// <summary>
        /// Gets the folders.
        /// </summary>
        /// <value>
        /// The folders.
        /// </value>
        public FolderShaman Folders { get; }

        /// <summary>
        /// Gets the sub site.
        /// </summary>
        /// <param name="site">The site.</param>
        /// <returns></returns>
        public SharePointFileManager GetSubSite(string site) => new SharePointFileManager(_context, site);

        /// <summary>
        /// Debugs the site asynchronous.
        /// </summary>
        public async Task DebugSiteAsync()
        {
            var site = _web;
            _context.Load(site, s => s.Title, s => s.Id, s => s.Language, s => s.UIVersion, s => s.CurrentUser, s => s.Description, s => s.Created, s => s.Webs);
            await _context.ExecuteQueryAsync();
            if (_log == null) return;
            _log.LogInformation($"SITE Title: {site.Title}  Description: {site.Description}");
            _log.LogInformation($"  Logged in as: {site.CurrentUser.LoginName}");
            _log.LogInformation($"  Available sites: {site.Webs.Count}  Title: {site.CurrentUser.Title}");
            for (var i = 0; i < site.Webs.Count; i++)
                _log.LogInformation($"    Sub Site: Url: {site.Webs[i].Url} Title: {site.Webs[i].Title} Description: {site.Webs[i].Description})");
        }
    }
}
