namespace SharePoint.IO
{
    /// <summary>
    /// ISharePointConnectionString
    /// </summary>
    public interface ISharePointConnectionString
    {
        string this[string name] { get; }
    }
}
