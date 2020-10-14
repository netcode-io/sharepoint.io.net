using Microsoft.SharePoint.Client;
using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using File = Microsoft.SharePoint.Client.File;

namespace SharePoint.IO.Managers
{
    public class FileShaman
    {
        readonly Web _web;
        public static char[] TrimChars = { '/' };
        const string StyleLibraryUrl = "Style Library";

        public FileShaman(Web web) => _web = web;

        public Lazy<FolderShaman> Folder => new Lazy<FolderShaman>(() => new FolderShaman(_web));

        public async Task UploadToStyleLibraryAsync(string path, string serverFolder = null)
        {
            var file = path.Replace("\\", "/");
            var libraryUrl = $"{_web.ServerRelativeUrl.TrimEnd(TrimChars)}/{StyleLibraryUrl}/";
            Console.WriteLine($"Uploading file {file} to {libraryUrl}{serverFolder}");
            await Folder.Value.EnsureFoldersAsync(libraryUrl, serverFolder, file);
            await CheckOutFileAsync(file, libraryUrl, serverFolder);
            var uploadedFile = await AddFileToAsync(libraryUrl, path, file, serverFolder);
            await CheckInPublishAndApproveFileAsync(uploadedFile);
        }

        public Task<File> AddFileToAsync(string libraryUrl, string path, string file = null, string serverFolder = null, string[] defines = null) => AddFileToAsync(libraryUrl, GetContent(path, defines), file ?? Path.GetFileName(path), serverFolder);
        public Task<File> AddFileToAsync(string libraryUrl, Stream stream, string file, string serverFolder = null, string[] defines = null) => AddFileToAsync(libraryUrl, GetContent(stream, defines), file, serverFolder);
        public async Task<File> AddFileToAsync(string libraryUrl, byte[] content, string file, string serverFolder)
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
            var uploadedFile = folder.Files.Add(spFile);
            _web.Context.Load(uploadedFile, f => f.CheckOutType, f => f.Level);
            await _web.Context.ExecuteQueryAsync();
            return uploadedFile;
        }

        byte[] GetContent(string path, string[] defines) =>
            defines == null
                ? System.IO.File.ReadAllBytes(path)
                : Encoding.UTF8.GetBytes(PreProcess.Process(System.IO.File.ReadAllText(path), defines));

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
                : Encoding.UTF8.GetBytes(PreProcess.Process(Encoding.UTF8.GetString(bytes), defines));
        }

        public async Task CheckInPublishAndApproveFileAsync(File file)
        {
            if (file.CheckOutType != CheckOutType.None)
                file.CheckIn("Updating file", CheckinType.MajorCheckIn);
            if (file.Level == FileLevel.Draft)
                file.Publish("Updating file");
            file.Context.Load(file, f => f.ListItemAllFields);
            await file.Context.ExecuteQueryAsync();
            if (file.ListItemAllFields["_ModerationStatus"].ToString() == "2") //: pending
            {
                file.Approve("Updating file");
                await file.Context.ExecuteQueryAsync();
            }
        }

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
