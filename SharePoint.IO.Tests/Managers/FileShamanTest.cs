using Microsoft.Extensions.DependencyInjection;
using System.IO;
using System.Text;
using Xunit;

namespace SharePoint.IO.Managers
{
    public class FileShamanTest
    {
        [Fact]
        public void AddFile()
        {
            var context = Config.Services.GetRequiredService<ISharePointContext>();
            using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(@"
Simple File
#ifdef alpha
    DOES NOT REPLACE
#endif
")))
            using (var sharePoint = context.Connect())
            {
                stream.Position = 0;
                sharePoint.GetFileManager().Files.AddFileAsync("Shared Documents", stream, "Simple File.txt");
            }
        }

        [Fact]
        public void AddFileWithDefines()
        {
            var context = Config.Services.GetRequiredService<ISharePointContext>();
            var defines = new[] { "prod" };
            using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(@"
PreParse File
#ifdef alpha
    ALPHA
#endif
#ifdef prod
    PROD
#endif
")))
            using (var sharePoint = context.Connect())
            {
                stream.Position = 0;
                sharePoint.GetFileManager().Files.AddFileAsync("Shared Documents", stream, "PreParse File.txt", defines: defines);
            }
        }
    }
}
