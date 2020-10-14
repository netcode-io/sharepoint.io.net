using System.Threading.Tasks;

namespace SharePoint.IO.Managers
{
    internal static class ManagerExtensions
    {
        public static async Task SetBingMapsKeyAsync(this SPWebManager source, string key) =>
            await source.SetWebPropertiesAsync(false, (properties, _) =>
            {
                properties["BING_MAPS_KEY"] = key;
            });

        public static async Task SetViewPortMetaTagAsync(this SPWebManager source) =>
            await source.SetWebPropertiesAsync(true, (properties, _) =>
            {
                const string viewPortMetaTag = "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1, maximum-scale=1\" />";
                properties["seoincludecustommetatagpropertyname"] = "True";
                properties["seocustommetatagpropertyname"] = viewPortMetaTag;
            }, true);
    }
}
