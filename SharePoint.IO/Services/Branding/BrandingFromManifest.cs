using System;
using System.Collections.Generic;
using System.Reflection;

namespace SharePoint.IO.Services.Branding
{
    public class BrandingFromManifest : IBranding
    {
        public BrandingFromManifest(string target, IList<Assembly> assemblies) => throw new NotImplementedException();

        public string LogoUrl { get; }
        public string CssUrl { get; }
        public IEnumerable<string> Files { get; }
        public IEnumerable<string> Xslts { get; }
        public IEnumerable<string> Scripts { get; }
        public IEnumerable<DisplayTemplateMeta> DisplayTemplates { get; }
        public IEnumerable<PageLayoutMeta> PageLayouts { get; }
    }
}
