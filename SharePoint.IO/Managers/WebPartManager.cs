using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using System.Threading.Tasks;

namespace SharePoint.IO.Managers
{
    /// <summary>
    /// WebPartManager
    /// </summary>
    public class WebPartManager
    {
        readonly ClientContext _ctx;
        readonly ILogger _log;

        /// <summary>
        /// Initializes a new instance of the <see cref="WebPartManager" /> class.
        /// </summary>
        /// <param name="ctx">The client context.</param>
        /// <param name="page">The page.</param>
        /// <param name="log">The log.</param>
        public WebPartManager(ClientContext ctx, File page, ILogger log)
        {
            _ctx = ctx;
            CurrentPage = page;
            _log = log;
        }

        /// <summary>
        /// Gets the current page.
        /// </summary>
        /// <value>
        /// The current page.
        /// </value>
        public File CurrentPage { get; }

        /// <summary>
        /// Adds the web part to layout asynchronous.
        /// </summary>
        /// <param name="webpartPath">The webpart path.</param>
        /// <param name="zone">The zone.</param>
        /// <param name="order">The order.</param>
        /// <returns></returns>
        public Task<string> AddWebPartToLayoutAsync(string webpartPath, string zone, int order) => AddWebPartAsync(System.IO.File.ReadAllText(webpartPath), zone, order);
        /// <summary>
        /// Adds the web part asynchronous.
        /// </summary>
        /// <param name="xml">The XML.</param>
        /// <param name="zone">The zone.</param>
        /// <param name="order">The order.</param>
        /// <returns></returns>
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

        /// <summary>
        /// Deletes all asynchronous.
        /// </summary>
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
