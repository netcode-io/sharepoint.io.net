using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Publishing.Navigation;
using Microsoft.SharePoint.Client.Taxonomy;
using System.Threading.Tasks;

namespace SharePoint.IO.Managers
{
    /// <summary>
    /// NavigationType
    /// </summary>
    public enum NavigationType
    {
        /// <summary>
        /// The no sites or pages
        /// </summary>
        NoSitesOrPages = 0,
        /// <summary>
        /// The sub sites only
        /// </summary>
        SubSitesOnly = 1,
        /// <summary>
        /// The pages only
        /// </summary>
        PagesOnly = 2,
        /// <summary>
        /// The subsites and pages
        /// </summary>
        SubsitesAndPages = 3
    }

    /// <summary>
    /// TemplateCodes
    /// </summary>
    public struct TemplateCodes
    {
        /// <summary>
        /// The team site
        /// </summary>
        public const string TeamSite = "STS#0";
        /// <summary>
        /// The blank site
        /// </summary>
        public const string BlankSite = "STS#1";
        /// <summary>
        /// The wiki site
        /// </summary>
        public const string WikiSite = "WIKI#0";
        /// <summary>
        /// The blog site
        /// </summary>
        public const string BlogSite = "BLOG#0";
        /// <summary>
        /// The dev site
        /// </summary>
        public const string DevSite = "DEV#0";
        /// <summary>
        /// The publishing site
        /// </summary>
        public const string PublishingSite = "BLANKINTERNET#0";
        /// <summary>
        /// The enterprise search site
        /// </summary>
        public const string EnterpriseSearchSite = "SRCHCEN#0";
    }

    /// <summary>
    /// WebProvisionerManager
    /// </summary>
    public class WebProvisionerManager
    {
        readonly Web _web;
        readonly ClientContext _ctx;
        readonly ILogger _log;

        /// <summary>
        /// Initializes a new instance of the <see cref="WebProvisionerManager" /> class.
        /// </summary>
        /// <param name="web">The web.</param>
        /// <param name="ctx">The client context.</param>
        /// <param name="log">The log.</param>
        public WebProvisionerManager(Web web, ClientContext ctx, ILogger log)
        {
            _web = web;
            _ctx = ctx;
            _log = log;
        }

        /// <summary>
        /// Creates the team site asynchronous.
        /// </summary>
        /// <param name="relativeUrl">The relative URL.</param>
        /// <param name="title">The title.</param>
        /// <returns></returns>
        public Task<Web> CreateTeamSiteAsync(string relativeUrl, string title) => CreateWebAsync(relativeUrl, title, TemplateCodes.TeamSite);
        /// <summary>
        /// Creates the publishing site asynchronous.
        /// </summary>
        /// <param name="relativeUrl">The relative URL.</param>
        /// <param name="title">The title.</param>
        /// <returns></returns>
        public Task<Web> CreatePublishingSiteAsync(string relativeUrl, string title) => CreateWebAsync(relativeUrl, title, TemplateCodes.PublishingSite);
        /// <summary>
        /// Creates the web asynchronous.
        /// </summary>
        /// <param name="relativeUrl">The relative URL.</param>
        /// <param name="title">The title.</param>
        /// <param name="templateCode">The template code.</param>
        /// <returns></returns>
        public async Task<Web> CreateWebAsync(string relativeUrl, string title, string templateCode)
        {
            var web = _web.Webs.Add(new WebCreationInformation
            {
                Url = relativeUrl,
                Title = title,
                UseSamePermissionsAsParentSite = true,
                WebTemplate = templateCode,
                Language = 1033
            });
            await _ctx.ExecuteQueryAsync();
            return web;
        }

        /// <summary>
        /// Sets the welcome page asynchronous.
        /// </summary>
        /// <param name="pageRelativeUrl">The page relative URL.</param>
        /// <param name="web">The web.</param>
        public async Task SetWelcomePageAsync(string pageRelativeUrl, Web web = null)
        {
            var rootFolder = (web ?? _web).RootFolder;
            rootFolder.WelcomePage = pageRelativeUrl;
            rootFolder.Update();
            await _ctx.ExecuteQueryAsync();
        }

        /// <summary>
        /// Sets the current navigation asynchronous.
        /// </summary>
        /// <param name="web">The web.</param>
        /// <param name="source">The source.</param>
        /// <param name="displaySettings">The display settings.</param>
        public async Task SetCurrentNavigationAsync(Web web, StandardNavigationSource source, NavigationType displaySettings)
        {
            var targetWeb = web ?? _web;
            var ctx = targetWeb.Context;
            var settings = new WebNavigationSettings(ctx, targetWeb);
            settings.CurrentNavigation.Source = source;
            settings.Update(TaxonomySession.GetTaxonomySession(ctx));
            ctx.Load(targetWeb, w => w.AllProperties);
            await ctx.ExecuteQueryAsync();
            targetWeb.AllProperties["__CurrentNavigationIncludeTypes"] = ((int)displaySettings).ToString();
            targetWeb.AllProperties["__NavigationShowSiblings"] = bool.TrueString;
            targetWeb.Update();
            ctx.Load(targetWeb);
            await ctx.ExecuteQueryAsync();
        }
    }
}

