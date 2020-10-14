using Microsoft.SharePoint.Client;
using System;
using System.Threading.Tasks;

namespace SharePoint.IO.Managers
{
    public class SPWebManager
    {
        readonly ClientContext _ctx;

        public SPWebManager(ClientContext ctx) => _ctx = ctx;

        public Web CurrentWeb { get; private set; }

        public Site CurrentSite { get; private set; }

        public Task<Web> LoadWebAsync(string url = null) => LoadWebAsync(string.IsNullOrEmpty(url) ? null : _ctx.Site.OpenWeb(url));

        public async Task<Web> LoadWebAsync(Web web)
        {
            CurrentWeb = web;
            _ctx.Load(CurrentWeb);
            await _ctx.ExecuteQueryAsync();
            return CurrentWeb;
        }

        public async Task<Site> LoadSiteAsync(Site site = null)
        {
            CurrentSite = site ?? _ctx.Site;
            _ctx.Load(CurrentSite);
            await _ctx.ExecuteQueryAsync();
            return CurrentSite;
        }

        public async Task ExecuteWebQueryAsync(Web web = null)
        {
            var targetWeb = web ?? CurrentWeb ?? _ctx.Web;
            targetWeb.Update();
            _ctx.Load(targetWeb);
            await _ctx.ExecuteQueryAsync();
        }

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
