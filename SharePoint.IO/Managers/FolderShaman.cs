using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Security.Cryptography.X509Certificates;
using System.Threading;
using System.Threading.Tasks;

namespace SharePoint.IO.Managers
{
    /// <summary>
    /// FolderShaman
    /// </summary>
    public class FolderShaman
    {
        readonly Web _web;
        readonly ILogger _log;

        /// <summary>
        /// Initializes a new instance of the <see cref="FolderShaman" /> class.
        /// </summary>
        /// <param name="web">The web.</param>
        /// <param name="log">The log.</param>
        public FolderShaman(Web web, ILogger log)
        {
            _web = web;
            _log = log;
        }

        /// <summary>
        /// Ensures the path asynchronous.
        /// </summary>
        /// <param name="listUrl">The list URL.</param>
        /// <param name="folderUrl">The folder URL.</param>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="endsWithFile">if set to <c>true</c> [ends with file].</param>
        public async Task EnsurePathAsync(string listUrl, string folderUrl, string fileName, bool endsWithFile = true)
        {
            if (string.IsNullOrEmpty(listUrl))
                throw new ArgumentNullException(nameof(listUrl));
            if (string.IsNullOrEmpty(folderUrl))
                throw new ArgumentNullException(nameof(folderUrl));
            if (string.IsNullOrEmpty(fileName))
                throw new ArgumentNullException(nameof(fileName));
            var folder = await EnsureFolderAsync(listUrl, folderUrl);
            var fileUrls = fileName.Replace("\\", "/").Split(new[] { "/" }, StringSplitOptions.RemoveEmptyEntries);
            if (endsWithFile)
                fileUrls = fileUrls.Take(fileUrls.Count() - 1).ToArray();
            var parent = folder;
            foreach (var fileUrl in fileUrls)
                parent = await EnsureFolderAsync(listUrl, fileUrl, parent);
        }

        /// <summary>
        /// Ensures the folder asynchronous.
        /// </summary>
        /// <param name="listUrl">The list URL.</param>
        /// <param name="folderUrl">The folder URL.</param>
        /// <param name="parentFolder">The parent folder.</param>
        /// <param name="retrievals">The retrievals.</param>
        /// <returns></returns>
        public async Task<Folder> EnsureFolderAsync(string listUrl, string folderUrl, Folder parentFolder = null, params Expression<Func<Folder, object>>[] retrievals)
        {
            if (string.IsNullOrEmpty(listUrl))
                throw new ArgumentNullException(nameof(listUrl));
            if (string.IsNullOrEmpty(folderUrl))
                throw new ArgumentNullException(nameof(folderUrl));
            listUrl = listUrl.Replace("\\", "/");
            folderUrl = folderUrl.Replace("\\", "/");
            Folder folder;
            var list = await GetListFromUrlAsync(listUrl);
            var scope = new ExceptionHandlingScope(_web.Context);
            using (scope.StartScope())
            {
                using (scope.StartTry())
                    GetExistingFolder(listUrl, folderUrl, parentFolder, null);
                using (scope.StartCatch())
                    CreateFolder(list, folderUrl, parentFolder);
                using (scope.StartFinally())
                    folder = GetExistingFolder(listUrl, folderUrl, parentFolder, retrievals);
            }
            var attempt = 0;
            while (true)
                try
                {
                    await _web.Context.ExecuteQueryAsync();
                    return folder;
                }
                catch (ServerException e) when (e.Message == "File Not Found." && attempt++ <= 2)
                {
                    _log?.LogInformation($"Retry{attempt}: {e.Message}");
                    Thread.Sleep(100);
                }
        }

        /// <summary>
        /// Ensures all folders asynchronous.
        /// </summary>
        /// <param name="listUrl">The list URL.</param>
        /// <param name="folderUrls">The folder urls.</param>
        /// <param name="batchSize">Size of the batch.</param>
        /// <returns></returns>
        /// <exception cref="Dictionary<string, Folder>"></exception>
        public async Task EnsureAllFoldersAsync(string listUrl, IEnumerable<string> folderUrls, int batchSize = 50)
        {
            if (string.IsNullOrEmpty(listUrl))
                throw new ArgumentNullException(nameof(listUrl));
            if (folderUrls == null)
                return;
            var levels = folderUrls.SelectMany(s =>
            {
                var ss = s.Replace("\\", "/").Split(new[] { "/" }, StringSplitOptions.RemoveEmptyEntries);
                return ss.Select((_, level) => (level, path: $"{listUrl}/{string.Join("/", ss.Take(level + 1))}"));
            }).Distinct().ToLookup(s => s.level, s => s.path).ToList();
            var list = await GetListFromUrlAsync(listUrl);
            var folders = new Dictionary<string, Folder> { { listUrl, null } };
            foreach (var level in levels)
                foreach (var batch in level.GroupAt(batchSize))
                {
                    foreach (var path in batch)
                    {
                        var scope = new ExceptionHandlingScope(_web.Context);
                        using (scope.StartScope())
                        {
                            var folderPath = Path.GetDirectoryName(path).Replace("\\", "/");
                            if (!folders.TryGetValue(folderPath, out var parentFolder))
                                throw new InvalidOperationException(folderPath);
                            var folderUrl = Path.GetFileName(path);
                            using (scope.StartTry())
                                GetExistingFolder(listUrl, folderUrl, parentFolder, null);
                            using (scope.StartCatch())
                                CreateFolder(list, folderUrl, parentFolder);
                            using (scope.StartFinally())
                                folders[path] = GetExistingFolder(listUrl, folderUrl, parentFolder, null);
                        }
                    }
                    var attempt = 0;
                    while (true)
                        try
                        {
                            await _web.Context.ExecuteQueryAsync();
                            break;
                        }
                        catch (ServerException e) when (e.Message == "File Not Found." && attempt++ <= 3)
                        {
                            _log?.LogInformation($"Retry{attempt}: {e.Message}");
                            Thread.Sleep(500);
                        }
                }
        }

        public async Task DeleteFolderAsync(string listUrl, string folderUrl, Folder parentFolder = null)
        {
            if (string.IsNullOrEmpty(listUrl))
                throw new ArgumentNullException(nameof(listUrl));
            if (string.IsNullOrEmpty(folderUrl))
                return;
            var scope = new ExceptionHandlingScope(_web.Context);
            using (scope.StartScope())
            {
                using (scope.StartTry())
                {
                    var folderPath = $"{listUrl.TrimEnd(FileShaman.TrimChars)}/{folderUrl}";
                    var folder = _web.GetFolderByServerRelativeUrl($"{listUrl.TrimEnd(FileShaman.TrimChars)}/{folderUrl}");
                    folder.DeleteObject();
                }
                using (scope.StartCatch()) { }
            }
            var attempt = 0;
            while (true)
                try
                {
                    await _web.Context.ExecuteQueryAsync();
                    return;
                }
                catch (ServerException e) when (e.Message == "File Not Found." && attempt++ <= 3)
                {
                    _log?.LogInformation($"Retry{attempt}: {e.Message}");
                    Thread.Sleep(500);
                }
        }

        /// <summary>
        /// Deletes all folders asynchronous.
        /// </summary>
        /// <param name="listUrl">The list URL.</param>
        /// <param name="folderUrls">The folder urls.</param>
        /// <param name="batchSize">Size of the batch.</param>
        /// <returns></returns>
        /// <exception cref="Dictionary<string, Folder>"></exception>
        public async Task DeleteAllFoldersAsync(string listUrl, IEnumerable<string> folderUrls, int batchSize = 50)
        {
            if (string.IsNullOrEmpty(listUrl))
                throw new ArgumentNullException(nameof(listUrl));
            if (folderUrls == null)
                return;
            var urls = folderUrls.Aggregate(new HashSet<string>(), (s, y) =>
            {
                var key = $"{y}/";
                foreach (var x in s.ToList())
                    if (x.StartsWith(key))
                        s.Remove(x);
                    else if (key.StartsWith(x))
                        return s;
                s.Add(key);
                return s;
            }).Distinct().Select(x => x.Substring(0, x.Length - 1)).ToList();
            foreach (var batch in urls.GroupAt(batchSize))
            {
                foreach (var path in batch)
                {
                    var scope = new ExceptionHandlingScope(_web.Context);
                    using (scope.StartScope())
                    {
                        using (scope.StartTry())
                        {
                            var folderPath = $"{listUrl.TrimEnd(FileShaman.TrimChars)}/{path}";
                            var folder = _web.GetFolderByServerRelativeUrl($"{listUrl.TrimEnd(FileShaman.TrimChars)}/{path}");
                            folder.DeleteObject();
                        }
                        using (scope.StartCatch()) { }
                    }
                }
                var attempt = 0;
                while (true)
                    try
                    {
                        await _web.Context.ExecuteQueryAsync();
                        break;
                    }
                    catch (ServerException e) when (e.Message == "File Not Found." && attempt++ <= 3)
                    {
                        _log?.LogInformation($"Retry{attempt}: {e.Message}");
                        Thread.Sleep(500);
                    }
            }
        }

        Folder GetExistingFolder(string listUrl, string folderUrl, Folder parentFolder, Expression<Func<Folder, object>>[] retrievals)
        {
            var folderServerRelativeUrl = GetFolderServerRelativeUrl(listUrl, folderUrl, parentFolder);
            var folder = _web.GetFolderByServerRelativeUrl(folderServerRelativeUrl);
            _web.Context.Load(folder);
            if (retrievals != null && retrievals.Length > 0)
                _web.Context.Load(folder, retrievals);
            return folder;
        }

        Folder CreateFolder(List list, string folderUrl, Folder parentFolder)
        {
            if (parentFolder == null)
                parentFolder = list.RootFolder;
            var folder = parentFolder.Folders.Add(folderUrl);
            _web.Context.Load(folder);
            return folder;
        }

        string GetFolderServerRelativeUrl(string listUrl, string folderUrl, Folder parentFolder)
        {
            var folderServerRelativeUrl = parentFolder == null
                ? $"{listUrl.TrimEnd(FileShaman.TrimChars)}/{folderUrl}"
                : $"{parentFolder.ServerRelativeUrl.TrimEnd(FileShaman.TrimChars)}/{folderUrl}";
            return folderServerRelativeUrl;
        }

        async Task<List> GetListFromUrlAsync(string listUrl)
        {
            var lists = _web.Lists;
            _web.Context.Load(_web);
            _web.Context.Load(lists, l => l.Include(ll => ll.DefaultViewUrl));
            await _web.Context.ExecuteQueryAsync();
            var list = lists.Where(l => l.DefaultViewUrl.IndexOf(listUrl, StringComparison.CurrentCultureIgnoreCase) >= 0).FirstOrDefault();
            return list;
        }

        /// <summary>
        /// Enumerates the directories.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <param name="searchPattern">The search pattern.</param>
        /// <param name="searchOption">The search option.</param>
        /// <returns></returns>
        /// <exception cref="System.ArgumentNullException">path</exception>
        /// <exception cref="System.NotSupportedException"></exception>
        public async Task<IEnumerable<string>> EnumerateDirectoriesAsync(Folder path, string searchPattern, SearchOption searchOption)
        {
            if (path == null)
                throw new ArgumentNullException(nameof(path));
            if (searchPattern != "*")
                throw new NotSupportedException();
            _web.Context.Load(path, s => s.Folders);
            await _web.Context.ExecuteQueryAsync();
            var rootPath = path.ServerRelativeUrl;
            var list = path.Folders.Select(s => $"{rootPath}/{s.Name}").ToList();
            if (searchOption == SearchOption.AllDirectories)
                foreach (var folder in path.Folders)
                    list.AddRange(await EnumerateDirectoriesAsync(folder, searchPattern, searchOption));
            return list;
        }

        /// <summary>
        /// Enumerates the files.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <param name="searchPattern">The search pattern.</param>
        /// <param name="searchOption">The search option.</param>
        /// <returns></returns>
        /// <exception cref="System.ArgumentNullException">path</exception>
        /// <exception cref="System.NotSupportedException"></exception>
        public async Task<IEnumerable<string>> EnumerateFilesAsync(Folder path, string searchPattern, SearchOption searchOption)
        {
            if (path == null)
                throw new ArgumentNullException(nameof(path));
            if (searchPattern != "*")
                throw new NotSupportedException();
            _web.Context.Load(path, s => s.Folders, s => s.Files);
            await _web.Context.ExecuteQueryAsync();
            var rootPath = path.ServerRelativeUrl;
            var list = path.Files.Select(s => $"{rootPath}/{s.Name}").ToList();
            if (searchOption == SearchOption.AllDirectories)
                foreach (var folder in path.Folders)
                    list.AddRange(await EnumerateFilesAsync(folder, searchPattern, searchOption));
            return list;
        }

        /// <summary>
        /// Enumerates the entities asynchronous.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <param name="searchPattern">The search pattern.</param>
        /// <param name="searchOption">The search option.</param>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException">
        /// nameof(path)
        /// or
        /// nameof(path)
        /// </exception>
        public async Task<IEnumerable<(char entity, string path)>> EnumerateEntitiesAsync(Folder path, string searchPattern, SearchOption searchOption)
        {
            if (path == null)
                throw new ArgumentNullException(nameof(path));
            if (searchPattern != "*")
                throw new NotSupportedException();
            _web.Context.Load(path, s => s.Folders, s => s.Files);
            await _web.Context.ExecuteQueryAsync();
            var rootPath = path.ServerRelativeUrl;
            var list = new List<(char entity, string path)>();
            list.AddRange(path.Folders.Select(s => $"{rootPath}/{s.Name}").ToList().Select(s => ('D', s)));
            list.AddRange(path.Files.Select(s => $"{rootPath}/{s.Name}").ToList().Select(s => ('F', s)));
            if (searchOption == SearchOption.AllDirectories)
                foreach (var folder in path.Folders)
                    list.AddRange(await EnumerateEntitiesAsync(folder, searchPattern, searchOption));
            return list;
        }
    }
}
