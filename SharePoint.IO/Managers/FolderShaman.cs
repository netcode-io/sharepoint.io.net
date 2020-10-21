﻿using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using System;
using System.Linq;
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
        /// <param name="filePath">The file path.</param>
        /// <param name="fileFolder">The file folder.</param>
        /// <param name="fileName">Name of the file.</param>
        public async Task EnsurePathAsync(string filePath, string fileFolder, string fileName)
        {
            var folder = await EnsureFolderAsync(filePath, fileFolder);
            var folderUrls = fileName.Split(new[] { "/" }, StringSplitOptions.RemoveEmptyEntries);
            folderUrls = folderUrls.Take(folderUrls.Count() - 1).ToArray();
            var parent = folder;
            foreach (var folderUrl in folderUrls)
                parent = await EnsureFolderAsync(filePath, folderUrl, parent);
        }

        /// <summary>
        /// Ensures the folder asynchronous.
        /// </summary>
        /// <param name="listUrl">The list URL.</param>
        /// <param name="folderUrl">The folder URL.</param>
        /// <param name="parentFolder">The parent folder.</param>
        /// <returns></returns>
        public async Task<Folder> EnsureFolderAsync(string listUrl, string folderUrl, Folder parentFolder = null)
        {
            if (string.IsNullOrEmpty(folderUrl))
                return null;
            Folder folder;
            var list = await GetListFromUrlAsync(listUrl);
            var scope = new ExceptionHandlingScope(_web.Context);
            using (scope.StartScope())
            {
                using (scope.StartTry())
                    GetExistingFolder(listUrl, folderUrl, parentFolder);
                using (scope.StartCatch())
                    CreateFolder(list, folderUrl, parentFolder);
                using (scope.StartFinally())
                    folder = GetExistingFolder(listUrl, folderUrl, parentFolder);
            }
            await _web.Context.ExecuteQueryAsync();
            return folder;
        }

        Folder GetExistingFolder(string listUrl, string folderUrl, Folder parentFolder)
        {
            var folderServerRelativeUrl = GetFolderServerRelativeUrl(listUrl, folderUrl, parentFolder);
            var folder = _web.GetFolderByServerRelativeUrl(folderServerRelativeUrl);
            _web.Context.Load(folder);
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
    }
}
