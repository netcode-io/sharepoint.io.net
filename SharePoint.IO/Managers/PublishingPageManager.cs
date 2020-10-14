using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Publishing;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SharePoint.IO.Managers
{
    public class PublishingPageManager
    {
        readonly Web _web;
        readonly ClientContext _ctx;

        public PublishingPageManager(Web web, ClientContext ctx)
        {
            _web = web;
            _ctx = ctx;
        }

        public File GetPage(string pageName) => _web.GetFileByServerRelativeUrl($"{_web.ServerRelativeUrl.TrimEnd('/')}/Pages/{pageName}");

        public Task<ListItem> CreatePageFromLayoutAsync(string layoutRelativeUrl, string pageNameWithExtension, string title) =>
            CreatePageFromLayoutWithMetaAsync(layoutRelativeUrl, pageNameWithExtension, new Dictionary<string, string>
            {
                { "Title", title }
            });

        public async Task<ListItem> CreatePageFromLayoutWithMetaAsync(string layoutRelativeUrl, string pageNameWithExtension, Dictionary<string, string> meta)
        {
            var page = await CreatePublishingPageAsync(layoutRelativeUrl, pageNameWithExtension);
            var pageItem = page.ListItem;
            SetMetadataForItem(pageItem, meta);
            pageItem.File.CheckIn("page created from code", CheckinType.MajorCheckIn);
            await _ctx.ExecuteQueryAsync();
            return pageItem;
        }

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
