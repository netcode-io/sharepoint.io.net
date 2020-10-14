using Microsoft.SharePoint.Client;
using SharePoint.IO.Managers;
using System;
using System.Threading.Tasks;

namespace SharePoint.IO
{
    public class SharePointFileManager
    {
        readonly ClientContext _context;
        readonly Web _web;

        public SharePointFileManager(ClientContext context, string url = null)
        {
            _context = context;
            _web = new SPWebManager(_context).LoadWebAsync(url).ConfigureAwait(false).GetAwaiter().GetResult();
            Files = new FileShaman(_web);
            Folders = new FolderShaman(_web);
        }

        public FileShaman Files { get; }

        public FolderShaman Folders { get; }

        public SharePointFileManager SubSite(string site) => new SharePointFileManager(_context, site);

        public async Task DebugSiteAsync()
        {
            var site = _web;
            _context.Load(site, s => s.Title, s => s.Id, s => s.Language, s => s.UIVersion, s => s.CurrentUser, s => s.Description, s => s.Created, s => s.Webs);
            await _context.ExecuteQueryAsync();
            Console.WriteLine($"SITE Title: {site.Title}  Description: {site.Description}");
            Console.WriteLine($"  Logged in as: {site.CurrentUser.LoginName}");
            Console.WriteLine($"  Available sites: {site.Webs.Count}  Title: {site.CurrentUser.Title}");
            for (var i = 0; i < site.Webs.Count; i++)
                Console.WriteLine($"    Sub Site: Url: {site.Webs[i].Url} Title: {site.Webs[i].Title} Description: {site.Webs[i].Description})");
        }
    }
}
