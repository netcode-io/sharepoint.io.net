using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Publishing;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SharePoint.IO.Managers
{
    /// <summary>
    /// 
    /// </summary>
    public class PublishingPageManager
    {
        readonly Web _web;
        readonly ClientContext _ctx;
        readonly ILogger _log;

        /// <summary>
        /// Initializes a new instance of the <see cref="PublishingPageManager"/> class.
        /// </summary>
        /// <param name="web">The web.</param>
        /// <param name="ctx">The client context.</param>
        /// <param name="log">The log.</param>
        public PublishingPageManager(Web web, ClientContext ctx, ILogger log)
        {
            _web = web;
            _ctx = ctx;
            _log = log;
        }

        /// <summary>
        /// Gets the page.
        /// </summary>
        /// <param name="pageName">Name of the page.</param>
        /// <returns></returns>
        public File GetPage(string pageName) => _web.GetFileByServerRelativeUrl($"{_web.ServerRelativeUrl.TrimEnd('/')}/Pages/{pageName}");

        /// <summary>
        /// Creates the page from layout asynchronous.
        /// </summary>
        /// <param name="layoutRelativeUrl">The layout relative URL.</param>
        /// <param name="pageNameWithExtension">The page name with extension.</param>
        /// <param name="title">The title.</param>
        /// <returns></returns>
        public Task<ListItem> CreatePageFromLayoutAsync(string layoutRelativeUrl, string pageNameWithExtension, string title) =>
            CreatePageFromLayoutWithMetaAsync(layoutRelativeUrl, pageNameWithExtension, new Dictionary<string, string>
            {
                { "Title", title }
            });

        /// <summary>
        /// Creates the page from layout with meta asynchronous.
        /// </summary>
        /// <param name="layoutRelativeUrl">The layout relative URL.</param>
        /// <param name="pageNameWithExtension">The page name with extension.</param>
        /// <param name="meta">The meta.</param>
        /// <returns></returns>
        public async Task<ListItem> CreatePageFromLayoutWithMetaAsync(string layoutRelativeUrl, string pageNameWithExtension, Dictionary<string, string> meta)
        {
            var page = await CreatePublishingPageAsync(layoutRelativeUrl, pageNameWithExtension);
            var pageItem = page.ListItem;
            SetMetadataForItem(pageItem, meta);
            pageItem.File.CheckIn("page created from code", CheckinType.MajorCheckIn);
            await _ctx.ExecuteQueryAsync();
            return pageItem;
        }

        /// <summary>
        /// Gets the page layout item asynchronous.
        /// </summary>
        /// <param name="layoutRelativeUrl">The layout relative URL.</param>
        /// <returns></returns>
        public async Task<ListItem> GetPageLayoutItemAsync(string layoutRelativeUrl)
        {
            _ctx.Load(_ctx.Site.RootWeb, w => w.ServerRelativeUrl);
            await _ctx.ExecuteQueryAsync();
            var pageLayout = _ctx.Site.RootWeb.GetFileByServerRelativeUrl($"{_ctx.Site.RootWeb.ServerRelativeUrl.TrimEnd('/')}/{layoutRelativeUrl}");
            var pageLayoutItem = pageLayout.ListItemAllFields;
            _ctx.Load(pageLayoutItem);
            await _ctx.ExecuteQueryAsync();
            return pageLayoutItem;
        }

        /// <summary>
        /// Creates the publishing page asynchronous.
        /// </summary>
        /// <param name="layoutRelativeUrl">The layout relative URL.</param>
        /// <param name="pageNameWithExtension">The page name with extension.</param>
        /// <returns></returns>
        public async Task<PublishingPage> CreatePublishingPageAsync(string layoutRelativeUrl, string pageNameWithExtension)
        {
            var layoutItem = await GetPageLayoutItemAsync(layoutRelativeUrl);
            var publishingWeb = PublishingWeb.GetPublishingWeb(_ctx, _web);
            var page = publishingWeb.AddPublishingPage(new PublishingPageInformation
            {
                Name = pageNameWithExtension,
                PageLayoutListItem = layoutItem
            });
            await _ctx.ExecuteQueryAsync();
            return page;
        }

        void SetMetadataForItem(ListItem listItem, Dictionary<string, string> meta)
        {
            foreach (var key in meta.Keys)
                listItem[key] = meta[key];
            listItem.Update();
        }
    }
}
