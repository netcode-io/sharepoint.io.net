using Microsoft.Extensions.Configuration;
using Microsoft.SharePoint.Client;
using SharePoint.IO.Managers;
using SharePoint.IO.Services.Branding;
using System;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace SharePoint.IO.Services
{
    public class BrandingService
    {
        readonly ClientContext _ctx;
        readonly string _target;
        readonly string[] _defines;
        readonly IBranding _branding;
        readonly Web _web;
        readonly FileShaman _fileShaman;
        readonly PageShaman _pageShaman;
        readonly JsInjector _jsInjector;
        readonly SPWebManager _webManager;

        public BrandingService(ClientContext ctx, IConfigurationSection section, string target = null, string[] defines = null)
            : this(ctx, target, defines, new BrandingFromSection(section, target)) { }
        public BrandingService(ClientContext ctx, string target = null, string[] defines = null, params object[] objs)
            : this(ctx, target, defines, new BrandingFromManifest(target, objs.Select(x => x.GetType().Assembly).Distinct().ToList())) { }
        public BrandingService(ClientContext ctx, string target = null, string[] defines = null, params Assembly[] assemblies)
            : this(ctx, target, defines, new BrandingFromManifest(target, assemblies)) { }
        public BrandingService(ClientContext ctx, string target, string[] defines, IBranding branding)
        {
            _ctx = ctx;
            _target = target;
            _defines = defines;
            _branding = branding;
            _webManager = new SPWebManager(_ctx);
            _web = _webManager.LoadWebAsync().ConfigureAwait(false).GetAwaiter().GetResult();
            _fileShaman = new FileShaman(_web);
            _pageShaman = new PageShaman(_web);
            _jsInjector = new JsInjector(_web);
        }

        public async Task DeployAllAsync(string appFolder = "APP")
        {
            await UploadAssetsAsync();
            await UploadPageLayoutsAsync(appFolder);
            await UploadDisplayTemplatesAsync();
            await UploadXsltoStyleLibraryAsync(appFolder);
            Console.ForegroundColor = ConsoleColor.Green;
            await SetSiteLogoAsync();
            await SetAltenateCssAsync();
            await SetViewPortMetaTagAsync();
            await SetScriptAsync();
        }

        public Task SetViewPortMetaTagAsync() => _webManager.SetViewPortMetaTagAsync();

        public async Task SetScriptAsync(string scriptDescription = null, string scriptLocation = null, bool allSites = true)
        {
            if (_branding.Scripts == null)
                return;
            await _jsInjector.AddJsLinkAsync(_branding.Scripts, scriptDescription, scriptLocation, allSites);
        }

        public async Task UploadXsltoStyleLibraryAsync(string appFolder)
        {
            if (_branding.Xslts == null)
                return;
            foreach (var path in _branding.Xslts)
                await _fileShaman.UploadToStyleLibraryAsync(path, appFolder);
        }

        public async Task SetSiteLogoAsync()
        {
            if (string.IsNullOrEmpty(_branding.LogoUrl))
                return;
            _web.SiteLogoUrl = GetUrl(_branding.LogoUrl);
            await _webManager.ExecuteWebQueryAsync();
            //Console.WriteLine($"Logo set: {_web.SiteLogoUrl}");
        }

        public async Task SetAltenateCssAsync()
        {
            if (string.IsNullOrEmpty(_branding.CssUrl))
                return;
            _web.AlternateCssUrl = GetUrl(_branding.CssUrl);
            await _webManager.ExecuteWebQueryAsync();
            //Console.WriteLine($"CSS set: {_web.AlternateCssUrl}");
        }

        public async Task UploadAssetsAsync()
        {
            if (_branding.Files == null)
                return;
            foreach (var path in _branding.Files)
                await _fileShaman.UploadToStyleLibraryAsync(path);
        }

        public async Task UploadDisplayTemplatesAsync()
        {
            if (_branding.DisplayTemplates == null)
                return;
            foreach (var displayTemplate in _branding.DisplayTemplates)
                await _pageShaman.UploadDisplayTemplateAsync(displayTemplate.Path, displayTemplate.Title, _defines);
        }

        public async Task UploadPageLayoutsAsync(string appFolder)
        {
            if (_branding.PageLayouts == null)
                return;
            foreach (var pagelayout in _branding.PageLayouts)
                await _pageShaman.UploadPageLayoutAsync(pagelayout.Path, pagelayout.Title, appFolder, pagelayout.PublishingAssociatedContentType, _defines);
        }

        string GetUrl(string url) => url.StartsWith("http://", StringComparison.OrdinalIgnoreCase) || url.StartsWith("https://", StringComparison.OrdinalIgnoreCase)
            ? url
            : (_web.ServerRelativeUrl + url).Replace("//", "/");
    }
}
