using Microsoft.SharePoint.Client;

namespace SharePoint.IO
{
    /// <summary>
    /// SharePointExtensions
    /// </summary>
    public static class SharePointExtensions
    {
        /// <summary>
        /// Gets the file manager.
        /// </summary>
        /// <param name="source">The source.</param>
        /// <param name="url">The URL.</param>
        /// <returns></returns>
        public static SharePointFileManager GetFileManager(this ClientContext source, string url = null) => new SharePointFileManager(source, url);
    }
}
