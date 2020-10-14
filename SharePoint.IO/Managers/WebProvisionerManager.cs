using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Publishing.Navigation;
using Microsoft.SharePoint.Client.Taxonomy;
using System.Threading.Tasks;

namespace SharePoint.IO.Managers
{
    public enum NavigationType
    {
        NoSitesOrPages = 0,
        SubSitesOnly = 1,
        PagesOnly = 2,
        SubsitesAndPages = 3
    }

    public struct TemplateCodes
    {
        public const string TeamSite = "STS#0";
        public const string BlankSite = "STS#1";
        public const string WikiSite = "WIKI#0";
        public const string BlogSite = "BLOG#0";
        public const string DevSite = "DEV#0";
        public const string PublishingSite = "BLANKINTERNET#0";
        public const string EnterpriseSearchSite = "SRCHCEN#0";
    }

    public class WebProvisionerManager
    {
        readonly Web _web;
        readonly ClientContext _ctx;

        public WebProvisionerManager(Web web, ClientContext ctx)
        {
            _web = web;
            _ctx = ctx;
        }

        public Task<Web> CreateTeamSiteAsync(string relativeUrl, string title) => CreateWebAsync(relativeUrl, title, TemplateCodes.TeamSite);
        public Task<Web> CreatePublishingSiteAsync(string relativeUrl, string title) => CreateWebAsync(relativeUrl, title, TemplateCodes.PublishingSite);
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

        public async Task SetWelcomePageAsync(string pageRelativeUrl, Web web = null)
        {
            var rootFolder = (web ?? _web).RootFolder;
            rootFolder.WelcomePage = pageRelativeUrl;
            rootFolder.Update();
            await _ctx.ExecuteQueryAsync();
        }

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

