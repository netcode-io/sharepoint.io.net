using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using System.Threading.Tasks;

namespace SharePoint.IO.Managers
{
    public class WebPartManager
    {
        readonly ClientContext _ctx;

        public WebPartManager(ClientContext ctx, File page)
        {
            _ctx = ctx;
            CurrentPage = page;
        }

        public File CurrentPage { get; }

        public Task<string> AddWebPartToLayoutAsync(string webpartPath, string zone, int order) => AddWebPartAsync(System.IO.File.ReadAllText(webpartPath), zone, order);
        public async Task<string> AddWebPartAsync(string xml, string zone, int order)
        {
            await CheckoutPageAsync();
            var wpmgr = CurrentPage.GetLimitedWebPartManager(PersonalizationScope.Shared);
            var definition = wpmgr.ImportWebPart(xml);
            wpmgr.AddWebPart(definition.WebPart, zone, order);
            CurrentPage.CheckIn(string.Empty, CheckinType.MajorCheckIn);
            await _ctx.ExecuteQueryAsync();
            return xml;
        }

        public async Task DeleteAllAsync()
        {
            await CheckoutPageAsync();
            var wpMgr = CurrentPage.GetLimitedWebPartManager(PersonalizationScope.Shared);
            var webparts = wpMgr.WebParts;
            _ctx.Load(webparts);
            await _ctx.ExecuteQueryAsync();
            foreach (var webpart in webparts)
                webpart.DeleteWebPart();
            CurrentPage.CheckIn("All webparts deleted by code", CheckinType.MajorCheckIn);
            await _ctx.ExecuteQueryAsync();
        }

        async Task CheckoutPageAsync()
        {
            _ctx.Load(CurrentPage);
            await _ctx.ExecuteQueryAsync();
            if (CurrentPage.CheckOutType == CheckOutType.None)
                CurrentPage.CheckOut();
        }
    }
}
