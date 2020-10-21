#if Tenant
using Microsoft.Online.SharePoint.TenantAdministration;
using Microsoft.SharePoint.Client;
using System.Threading;
using System.Threading.Tasks;

namespace SharePoint.IO.Managers
{
    /// <summary>
    /// TenantManager
    /// </summary>
    public class TenantManager
    {
        readonly ClientContext _ctx;
        readonly Tenant _tenant;

        /// <summary>
        /// Initializes a new instance of the <see cref="TenantManager"/> class.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        public TenantManager(ClientContext ctx)
        {
            _ctx = ctx;
            _tenant = new Tenant(_ctx);
        }

        /// <summary>
        /// Disables the deny add and customize pages.
        /// </summary>
        /// <param name="siteUrl">The site URL.</param>
        public async Task DisableDenyAddAndCustomizePages(string siteUrl)
        {
            var siteProperties = _tenant.GetSitePropertiesByUrl(siteUrl, true);
            _ctx.Load(siteProperties);
            await _ctx.ExecuteQueryAsync();

            siteProperties.DenyAddAndCustomizePages = DenyAddAndCustomizePagesStatus.Disabled;
            var result = siteProperties.Update();
            _ctx.Load(result);
            await _ctx.ExecuteQueryAsync();
            while (!result.IsComplete)
            {
                Thread.Sleep(result.PollingInterval);
                _ctx.Load(result);
                await _ctx.ExecuteQueryAsync();
            }
        }
    }
}
#endif