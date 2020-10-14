using System.Collections.Generic;

namespace SharePoint.IO.Services.Branding
{
    public interface IBranding
    {
        string LogoUrl { get; }
        string CssUrl { get; }
        IEnumerable<string> Files { get; }
        IEnumerable<string> Xslts { get; }
        IEnumerable<string> Scripts { get; }
        IEnumerable<DisplayTemplateMeta> DisplayTemplates { get; }
        IEnumerable<PageLayoutMeta> PageLayouts { get; }
    }
}
