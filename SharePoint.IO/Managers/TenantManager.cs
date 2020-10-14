//using Microsoft.Online.SharePoint.TenantAdministration;
//using Microsoft.SharePoint.Client;
//using System.Threading;
//using System.Threading.Tasks;

//namespace SharePoint.IO.Managers
//{
//    public class TenantManager
//    {
//        readonly ClientContext _ctx;
//        readonly Tenant _tenant;

//        public TenantManager(ClientContext ctx)
//        {
//            _ctx = ctx;
//            _tenant = new Tenant(_ctx);
//        }

//        public async Task DisableDenyAddAndCustomizePages(string siteUrl)
//        {
//            var siteProperties = _tenant.GetSitePropertiesByUrl(siteUrl, true);
//            _ctx.Load(siteProperties);
//            await _ctx.ExecuteQueryAsync();

//            siteProperties.DenyAddAndCustomizePages = DenyAddAndCustomizePagesStatus.Disabled;
//            var result = siteProperties.Update();
//            _ctx.Load(result);
//            await _ctx.ExecuteQueryAsync();
//            while (!result.IsComplete)
//            {
//                Thread.Sleep(result.PollingInterval);
//                _ctx.Load(result);
//                await _ctx.ExecuteQueryAsync();
//            }
//        }
//    }
//}