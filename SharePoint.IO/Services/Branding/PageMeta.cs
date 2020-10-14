namespace SharePoint.IO.Services.Branding
{
    public class PageMeta
    {
        public string Title { get; set; }
        //public string Description { get; set; }
        public string Path { get; set; }
    }

    public class PageLayoutMeta : PageMeta
    {
        public string PublishingAssociatedContentType { get; set; }
    }
    
    public class DisplayTemplateMeta : PageMeta { }
}
