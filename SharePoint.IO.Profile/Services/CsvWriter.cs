using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace SharePoint.IO.Profile.Services
{
    /// <summary>
    /// Stores single row within the CSV file
    /// </summary>
    public class CsvRow : List<string>
    {
        public string Entry { get; set; }
    }

    /// <summary>
    /// Processes the output CSV
    /// </summary>
    public class CsvWriter : StreamWriter
    {
        readonly ILogger _log;

        /// <summary>
        /// Initializes a new instance of the <see cref="CsvWriter"/> class.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <param name="log">The log.</param>
        public CsvWriter(string path, ILogger log) : base(path) => _log = log;

        /// <summary>
        /// Executes the specified CSV writer.
        /// </summary>
        /// <param name="row">The user entry instance.</param>
        /// <param name="log">The log.</param>
        /// <exception cref="System.ArgumentNullException">If the userdata instance is null</exception>
        public void CsvWrite(CsvRow row)
        {
            try
            {
                var b = new StringBuilder();
                var firstEntry = true;
                foreach (var cell in row)
                {
                    if (!firstEntry) b.Append(',');
                    if (cell.IndexOfAny(new char[] { '"', ',' }) != -1) b.Append($"\"{cell.Replace("\"", "\"\"")}\"");
                    else b.Append(cell);
                    firstEntry = false;
                }
                row.Entry = b.ToString();
                // Replaces a line break ("\r\n") with a single space.  Line breaks are typically encountered in the streetAddress field in Active Directory.
                if (row.Entry.Contains("\r\n")) row.Entry = row.Entry.Replace("\r\n", " ");
                WriteLine(row.Entry);
            }
            catch (Exception e) { _log?.LogCritical(e, string.Empty); }
        }
    }
}
