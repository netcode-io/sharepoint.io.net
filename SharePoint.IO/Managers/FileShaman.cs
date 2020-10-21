using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using SharePoint.IO.Services;
using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using File = Microsoft.SharePoint.Client.File;

namespace SharePoint.IO.Managers
{
    /// <summary>
    /// FileShaman
    /// </summary>
    public class FileShaman
    {
        public static char[] TrimChars = { '/' };
        const string StyleLibraryUrl = "Style Library";
        readonly Web _web;
        readonly ILogger _log;

        /// <summary>
        /// Initializes a new instance of the <see cref="FileShaman" /> class.
        /// </summary>
        /// <param name="web">The web.</param>
        /// <param name="log">The log.</param>
        public FileShaman(Web web, ILogger log)
        {
            _web = web;
            _log = log;
        }

        /// <summary>
        /// Gets the folder.
        /// </summary>
        /// <value>
        /// The folder.
        /// </value>
        public Lazy<FolderShaman> Folder => new Lazy<FolderShaman>(() => new FolderShaman(_web, _log));

        /// <summary>
        /// Uploads to style library asynchronous.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <param name="serverFolder">The server folder.</param>
        public async Task UploadToStyleLibraryAsync(string path, string serverFolder = null)
        {
            var file = path.Replace("\\", "/");
            var libraryUrl = $"{_web.ServerRelativeUrl.TrimEnd(TrimChars)}/{StyleLibraryUrl}/";
            _log?.LogInformation($"Uploading file {file} to {libraryUrl}{serverFolder}");
            await Folder.Value.EnsurePathAsync(libraryUrl, serverFolder, file);
            await CheckOutFileAsync(file, libraryUrl, serverFolder);
            var uploadedFile = await AddFileAsync(libraryUrl, path, file, serverFolder);
            await CheckInPublishAndApproveFileAsync(uploadedFile);
        }

        /// <summary>
        /// Adds the file asynchronous.
        /// </summary>
        /// <param name="libraryUrl">The library URL.</param>
        /// <param name="path">The path.</param>
        /// <param name="file">The file.</param>
        /// <param name="serverFolder">The server folder.</param>
        /// <param name="defines">The defines.</param>
        /// <returns></returns>
        public Task<File> AddFileAsync(string libraryUrl, string path, string file = null, string serverFolder = null, string[] defines = null) => AddFileAsync(libraryUrl, GetContent(path, defines), file ?? Path.GetFileName(path), serverFolder);
        /// <summary>
        /// Adds the file asynchronous.
        /// </summary>
        /// <param name="libraryUrl">The library URL.</param>
        /// <param name="stream">The stream.</param>
        /// <param name="file">The file.</param>
        /// <param name="serverFolder">The server folder.</param>
        /// <param name="defines">The defines.</param>
        /// <returns></returns>
        public Task<File> AddFileAsync(string libraryUrl, Stream stream, string file, string serverFolder = null, string[] defines = null) => AddFileAsync(libraryUrl, GetContent(stream, defines), file, serverFolder);
        /// <summary>
        /// Adds the file asynchronous.
        /// </summary>
        /// <param name="libraryUrl">The library URL.</param>
        /// <param name="content">The content.</param>
        /// <param name="file">The file.</param>
        /// <param name="serverFolder">The server folder.</param>
        /// <returns></returns>
        public async Task<File> AddFileAsync(string libraryUrl, byte[] content, string file, string serverFolder)
        {
            var serverUrl = $"{libraryUrl.EnsureEndsWith("/")}{serverFolder}";
            var fileUrl = $"{serverUrl.EnsureEndsWith("/")}{file}";
            var folder = _web.GetFolderByServerRelativeUrl(serverUrl);
            var spFile = new FileCreationInformation
            {
                Content = content,
                Url = fileUrl,
                Overwrite = true
            };
            _log?.LogInformation($"Adding file {file} to {libraryUrl}{serverFolder}");
            var uploadedFile = folder.Files.Add(spFile);
            _web.Context.Load(uploadedFile, f => f.CheckOutType, f => f.Level);
            await _web.Context.ExecuteQueryAsync();
            return uploadedFile;
        }

        byte[] GetContent(string path, string[] defines) =>
            defines == null
                ? System.IO.File.ReadAllBytes(path)
                : Encoding.UTF8.GetBytes(ContentService.Process(System.IO.File.ReadAllText(path), defines));

        byte[] GetContent(Stream stream, string[] defines)
        {
            byte[] bytes;
            using (var s = new MemoryStream())
            {
                stream.CopyTo(s);
                if (stream.CanSeek)
                    stream.Position = 0;
                bytes = s.ToArray();
            }
            return defines == null
                ? bytes
                : Encoding.UTF8.GetBytes(ContentService.Process(Encoding.UTF8.GetString(bytes), defines));
        }

        /// <summary>
        /// Checks the in publish and approve file asynchronous.
        /// </summary>
        /// <param name="file">The file.</param>
        /// <param name="comment">The comment.</param>
        /// <param name="checkInType">Type of the check in.</param>
        public async Task CheckInPublishAndApproveFileAsync(File file, string comment = null, CheckinType? checkInType = null)
        {
            if (file.CheckOutType != CheckOutType.None)
                file.CheckIn(comment ?? "Updating file", checkInType ?? CheckinType.MajorCheckIn);
            if (file.Level == FileLevel.Draft)
                file.Publish(comment ?? "Updating file");
            file.Context.Load(file, f => f.ListItemAllFields);
            await file.Context.ExecuteQueryAsync();
            if (file.ListItemAllFields["_ModerationStatus"].ToString() == "2") //: pending
            {
                file.Approve(comment ?? "Updating file");
                await file.Context.ExecuteQueryAsync();
            }
        }

        /// <summary>
        /// Checks the out file asynchronous.
        /// </summary>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="filePath">The file path.</param>
        /// <param name="fileFolder">The file folder.</param>
        public async Task CheckOutFileAsync(string fileName, string filePath, string fileFolder)
        {
            var fileUrl = $"{filePath}{fileFolder}{(string.IsNullOrEmpty(fileFolder) ? string.Empty : "/")}{fileName}";
            var file = _web.GetFileByServerRelativeUrl(fileUrl);
            _web.Context.Load(file, f => f.Exists);
            await _web.Context.ExecuteQueryAsync();
            if (!file.Exists)
                return;

            _web.Context.Load(file, f => f.CheckOutType);
            await _web.Context.ExecuteQueryAsync();
            if (file.CheckOutType != CheckOutType.None)
                file.UndoCheckOut();
            file.CheckOut();
            await _web.Context.ExecuteQueryAsync();
        }
    }
}
