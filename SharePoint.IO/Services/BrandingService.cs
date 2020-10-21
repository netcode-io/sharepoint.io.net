using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using SharePoint.IO.Managers;
using SharePoint.IO.Services.Branding;
using System;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace SharePoint.IO.Services
{
    /// <summary>
    /// BrandingService
    /// </summary>
    public class BrandingService
    {
        readonly ClientContext _ctx;
        readonly string _target;
        readonly string[] _defines;
        readonly IBranding _branding;
        readonly ILogger _log;
        readonly Web _web;
        readonly FileShaman _fileShaman;
        readonly PageShaman _pageShaman;
        readonly JsInjector _jsInjector;
        readonly SPWebManager _webManager;

        /// <summary>
        /// Initializes a new instance of the <see cref="BrandingService"/> class.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="section">The section.</param>
        /// <param name="target">The target.</param>
        /// <param name="defines">The defines.</param>
        /// <param name="log">The log.</param>
        public BrandingService(ClientContext ctx, IConfigurationSection section, string target = null, string[] defines = null, ILogger log = null)
            : this(ctx, target, defines, log, new BrandingFromSection(section, target)) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="BrandingService"/> class.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="target">The target.</param>
        /// <param name="defines">The defines.</param>
        /// <param name="log">The log.</param>
        /// <param name="objs">The objs.</param>
        public BrandingService(ClientContext ctx, string target = null, string[] defines = null, ILogger log = null, params object[] objs)
            : this(ctx, target, defines, log, new BrandingFromManifest(target, objs.Select(x => x.GetType().Assembly).Distinct().ToList())) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="BrandingService"/> class.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="target">The target.</param>
        /// <param name="defines">The defines.</param>
        /// <param name="log">The log.</param>
        /// <param name="assemblies">The assemblies.</param>
        public BrandingService(ClientContext ctx, string target = null, string[] defines = null, ILogger log = null, params Assembly[] assemblies)
            : this(ctx, target, defines, log, new BrandingFromManifest(target, assemblies)) { }
        /// <summary>
        /// Initializes a new instance of the <see cref="BrandingService"/> class.
        /// </summary>
        /// <param name="ctx">The CTX.</param>
        /// <param name="target">The target.</param>
        /// <param name="defines">The defines.</param>
        /// <param name="log">The log.</param>
        /// <param name="branding">The branding.</param>
        public BrandingService(ClientContext ctx, string target, string[] defines, ILogger log, IBranding branding)
        {
            _ctx = ctx;
            _target = target;
            _defines = defines;
            _branding = branding;
            _log = log;
            _webManager = new SPWebManager(_ctx, log);
            _web = _webManager.LoadWebAsync().ConfigureAwait(false).GetAwaiter().GetResult();
            _fileShaman = new FileShaman(_web, log);
            _pageShaman = new PageShaman(_web, log);
            _jsInjector = new JsInjector(_web, log);
        }

        /// <summary>
        /// Deploys all asynchronous.
        /// </summary>
        /// <param name="appFolder">The application folder.</param>
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

        /// <summary>
        /// Sets the view port meta tag asynchronous.
        /// </summary>
        /// <returns></returns>
        public Task SetViewPortMetaTagAsync() => _webManager.SetViewPortMetaTagAsync();

        /// <summary>
        /// Sets the script asynchronous.
        /// </summary>
        /// <param name="scriptDescription">The script description.</param>
        /// <param name="scriptLocation">The script location.</param>
        /// <param name="allSites">if set to <c>true</c> [all sites].</param>
        public async Task SetScriptAsync(string scriptDescription = null, string scriptLocation = null, bool allSites = true)
        {
            if (_branding.Scripts == null)
                return;
            await _jsInjector.AddJsLinkAsync(_branding.Scripts, scriptDescription, scriptLocation, allSites);
        }

        /// <summary>
        /// Uploads the xslto style library asynchronous.
        /// </summary>
        /// <param name="appFolder">The application folder.</param>
        public async Task UploadXsltoStyleLibraryAsync(string appFolder)
        {
            if (_branding.Xslts == null)
                return;
            foreach (var path in _branding.Xslts)
                await _fileShaman.UploadToStyleLibraryAsync(path, appFolder);
        }

        /// <summary>
        /// Sets the site logo asynchronous.
        /// </summary>
        public async Task SetSiteLogoAsync()
        {
            if (string.IsNullOrEmpty(_branding.LogoUrl))
                return;
            _web.SiteLogoUrl = GetUrl(_branding.LogoUrl);
            await _webManager.ExecuteWebQueryAsync();
            _log?.LogInformation($"Logo set: {_web.SiteLogoUrl}");
        }

        /// <summary>
        /// Sets the altenate CSS asynchronous.
        /// </summary>
        public async Task SetAltenateCssAsync()
        {
            if (string.IsNullOrEmpty(_branding.CssUrl))
                return;
            _web.AlternateCssUrl = GetUrl(_branding.CssUrl);
            await _webManager.ExecuteWebQueryAsync();
            _log?.LogInformation($"CSS set: {_web.AlternateCssUrl}");
        }

        /// <summary>
        /// Uploads the assets asynchronous.
        /// </summary>
        public async Task UploadAssetsAsync()
        {
            if (_branding.Files == null)
                return;
            foreach (var path in _branding.Files)
                await _fileShaman.UploadToStyleLibraryAsync(path);
        }

        /// <summary>
        /// Uploads the display templates asynchronous.
        /// </summary>
        public async Task UploadDisplayTemplatesAsync()
        {
            if (_branding.DisplayTemplates == null)
                return;
            foreach (var displayTemplate in _branding.DisplayTemplates)
                await _pageShaman.UploadDisplayTemplateAsync(displayTemplate.Path, displayTemplate.Title, _defines);
        }

        /// <summary>
        /// Uploads the page layouts asynchronous.
        /// </summary>
        /// <param name="appFolder">The application folder.</param>
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
