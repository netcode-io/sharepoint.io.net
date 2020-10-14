namespace SharePoint.IO
{
    internal static class Extensions
    {
        public static string EnsureEndsWith(this string s, string prefix) => s.EndsWith(prefix) ? s : s + prefix;
    }
}
