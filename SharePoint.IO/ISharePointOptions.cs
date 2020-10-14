using System.Net;

namespace SharePoint.IO
{
    /// <summary>
    /// ISharePointOptions
    /// </summary>
    public interface ISharePointOptions
    {
        string Endpoint { get; }
        NetworkCredential ServiceLogin { get; }
    }
}
