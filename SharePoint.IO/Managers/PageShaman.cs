using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using System.Linq;
using System.Threading.Tasks;

namespace SharePoint.IO.Managers
{
    /// <summary>
    /// PageShaman
    /// </summary>
    public class PageShaman
    {
        readonly Web _web;
        readonly ClientRuntimeContext _ctx;
        readonly ILogger _log;
        readonly FolderShaman _folderShaman;
        readonly FileShaman _fileShaman;
        const string PageLayoutContentTypeId = "0x01010007FF3E057FA8AB4AA42FCB67B453FFC100E214EEE741181F4E9F7ACC43278EE811";
        const string DefaultPageCType = ";#Article Page;#0x010100C568DB52D9D0A14D9B2FDCC96666E9F2007948130EC3DB064584E219954237AF3900242457EFB8B24247815D688C526CD44D;#";

        /// <summary>
        /// Initializes a new instance of the <see cref="PageShaman"/> class.
        /// </summary>
        /// <param name="web">The web.</param>
        /// <param name="log">The log.</param>
        public PageShaman(Web web, ILogger log)
        {
            _web = web;
            _ctx = web.Context;
            _log = log;
            _folderShaman = new FolderShaman(web, log);
            _fileShaman = new FileShaman(web, log);
        }

        /// <summary>
        /// Uploads the display template asynchronous.
        /// </summary>
        /// <param name="locationPath">The location path.</param>
        /// <param name="title">The title.</param>
        /// <param name="defines">The defines.</param>
        public async Task UploadDisplayTemplateAsync(string locationPath, string title, string[] defines = null)
        {
            var catalogPath = await GetCatalogPathByIdAsync((int)ListTemplateType.MasterPageCatalog);
            var destUrl = locationPath.Replace("\\", "/");
            _log?.LogInformation($"Uploading display template {locationPath} to {catalogPath}");
            //
            await _fileShaman.CheckOutFileAsync(locationPath, catalogPath, string.Empty);
            var uploadedFile = await _fileShaman.AddFileAsync(catalogPath, locationPath, destUrl, string.Empty, defines: defines);
            await SetPageMetadataAsync(uploadedFile, title);
            await _fileShaman.CheckInPublishAndApproveFileAsync(uploadedFile);
        }

        /// <summary>
        /// Uploads the page layout asynchronous.
        /// </summary>
        /// <param name="locationPath">The location path.</param>
        /// <param name="title">The title.</param>
        /// <param name="appFolder">The application folder.</param>
        /// <param name="publishingAssociatedContentType">Type of the publishing associated content.</param>
        /// <param name="defines">The defines.</param>
        public async Task UploadPageLayoutAsync(string locationPath, string title, string appFolder, string publishingAssociatedContentType, string[] defines = null)
        {
            var catalogPath = await GetCatalogPathByIdAsync((int)ListTemplateType.MasterPageCatalog);
            var destUrl = locationPath.Replace("\\", "/");
            _log?.LogInformation($"Uploading page layout {locationPath} to {catalogPath}");
            //
            await _folderShaman.EnsurePathAsync(catalogPath, appFolder, destUrl);
            await _fileShaman.CheckOutFileAsync(locationPath, catalogPath, appFolder);
            var uploadFile = await _fileShaman.AddFileAsync(catalogPath, locationPath, destUrl, appFolder, defines: defines);
            await SetPageLayoutMetadataAsync(uploadFile, title, publishingAssociatedContentType ?? DefaultPageCType);
            await _fileShaman.CheckInPublishAndApproveFileAsync(uploadFile);
        }

        async Task<string> GetCatalogPathByIdAsync(int typeCatalog)
        {
            var gallery = _web.GetCatalog(typeCatalog);
            _ctx.Load(gallery, g => g.RootFolder.ServerRelativeUrl);
            await _ctx.ExecuteQueryAsync();
            return $"{gallery.RootFolder.ServerRelativeUrl.TrimEnd(FileShaman.TrimChars)}/";
        }

        async Task SetPageMetadataAsync(File uploadFile, string title)
        {
            var item = uploadFile.ListItemAllFields;
            _web.Context.Load(item);
            item["Title"] = title;
            item.Update();
            await _web.Context.ExecuteQueryAsync();
        }

        async Task SetPageLayoutMetadataAsync(File uploadFile, string title, string publishingAssociatedContentType)
        {
            var gallery = _web.GetCatalog((int)ListTemplateType.MasterPageCatalog);
            _web.Context.Load(gallery, g => g.ContentTypes);
            await _web.Context.ExecuteQueryAsync();
            //
            var contentTypeId = gallery.ContentTypes.FirstOrDefault(ct => ct.StringId.StartsWith(PageLayoutContentTypeId)).StringId;
            var item = uploadFile.ListItemAllFields;
            _web.Context.Load(item);
            item["ContentTypeId"] = contentTypeId;
            item["Title"] = title;
            item["PublishingAssociatedContentType"] = publishingAssociatedContentType;
            item.Update();
            await _web.Context.ExecuteQueryAsync();
        }
    }
}
