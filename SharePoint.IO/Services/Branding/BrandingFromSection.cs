using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SharePoint.IO.Services.Branding
{
    public class BrandingFromSection : IBranding
    {
        public BrandingFromSection(IConfigurationSection section, string target)
        {
            foreach (var child in section.GetChildren())
                switch (child.Key.ToLowerInvariant())
                {
                    case "logourl": LogoUrl = child.Value; break;
                    case "cssurl": CssUrl = child.Value; break;
                    case "scripts":
                        Scripts = child.GetChildren().Where(x => x.Key == target).Select(x => x.Value).ToList();
                        break;
                    case "files":
                        Files = child.GetChildren().Select(x => x.Value).ToList();
                        break;
                    case "displaytemplates":
                        DisplayTemplates = child.GetChildren().Select(x =>
                        {
                            var r = new DisplayTemplateMeta();
                            foreach (var child in x.GetChildren())
                                switch (child.Key.ToLowerInvariant())
                                {
                                    case "title": r.Title = child.Value; break;
                                    case "path": r.Path = child.Value; break;
                                }
                            return r;
                        }).ToList();
                        break;
                    case "pageLayouts":
                        PageLayouts = child.GetChildren().Select(x =>
                        {
                            var r = new PageLayoutMeta();
                            foreach (var child in x.GetChildren())
                                switch (child.Key.ToLowerInvariant())
                                {
                                    case "title": r.Title = child.Value; break;
                                    case "path": r.Path = child.Value; break;
                                    case "publishingassociatedcontenttype": r.PublishingAssociatedContentType = child.Value; break;
                                }
                            return r;
                        }).ToList();
                        break;
                    case "xslts":
                        Xslts = child.GetChildren().Select(x => x.Value).ToList();
                        break;
                }
        }

        public string LogoUrl { get; }
        public string CssUrl { get; }
        public IEnumerable<string> Files { get; }
        public IEnumerable<string> Xslts { get; }
        public IEnumerable<string> Scripts { get; }
        public IEnumerable<DisplayTemplateMeta> DisplayTemplates { get; }
        public IEnumerable<PageLayoutMeta> PageLayouts { get; }
    }
}
