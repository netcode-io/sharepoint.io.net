using Microsoft.Extensions.Logging;
using System;
using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;

namespace SharePoint.IO.Profile.Services
{
    /// <summary>
    /// Processes the input CSV
    /// </summary>
    public class CsvReader
    {
        readonly ILogger _log;

        /// <summary>
        /// Initializes a new instance of the <see cref="CsvReader"/> class.
        /// </summary>
        /// <param name="log">The log.</param>
        public CsvReader(ILogger log) => _log = log;

        /// <summary>
        /// The delimiter as string
        /// </summary>
        string _delimiterAsString;

        /// <summary>
        /// Gets the delimiter.
        /// </summary>
        /// <value>
        /// The delimiter.
        /// </value>
        char Delimiter => DelimiterAsString[0];

        /// <summary>
        /// Gets or sets the delimiter as string.
        /// </summary>
        /// <value>
        /// The delimiter as string.
        /// </value>
        string DelimiterAsString
        {
            get
            {
                if (_delimiterAsString == null)
                    try
                    {
                        _delimiterAsString = CultureInfo.CurrentCulture.TextInfo.ListSeparator;
                        if (_delimiterAsString.Length != 1)
                            _delimiterAsString = ",";
                    }
                    catch (Exception) { _delimiterAsString = ","; }
                return _delimiterAsString;
            }
            set => _delimiterAsString = value;
        }

        /// <summary>
        /// Executes the specified reader.
        /// </summary>
        /// <param name="reader">The reader instance.</param>
        /// <param name="action">The logic to execute.</param>
        /// <exception cref="System.ArgumentNullException">If the reader instance is null</exception>
        public void Execute(TextReader reader, Action<Collection<string>, ILogger> action)
        {
            if (reader == null)
                throw new ArgumentNullException(nameof(reader));
            var lineNum = -1;
            try
            {
                string line = null;
                while ((line = reader.ReadLine()) != null)
                {
                    lineNum++;
                    if (lineNum == 0)
                    {
                        var separator = new[] { DelimiterAsString };
                        var commaSeperator = new[] { "," };
                        var length = line.Split(separator, StringSplitOptions.None).Length;
                        if (line.Split(commaSeperator, StringSplitOptions.None).Length > length)
                            DelimiterAsString = commaSeperator[0];
                    }
                    else if (lineNum <= 1 || !string.IsNullOrEmpty(line.Trim()))
                    {
                        try
                        {
                            var entries = ParseLineIntoEntries(line);
                            action.Invoke(entries, _log);
                        }
                        catch (Exception e) { _log?.LogCritical(e, string.Empty); }
                    }
                }
            }
            catch (Exception e) { _log?.LogCritical(e, string.Empty); }
        }

        /// <summary>
        /// Parses the line into entries.
        /// </summary>
        /// <param name="line">The line to parse.</param>
        /// <returns>A collection of columns</returns>
        Collection<string> ParseLineIntoEntries(string line)
        {
            var list = new Collection<string>();
            var lineArray = line.ToCharArray();
            var str = string.Empty;
            var flag = false;
            for (var i = 0; i < line.Length; i++)
            {
                if (!flag && string.IsNullOrEmpty(str))
                {
                    if (char.IsWhiteSpace(lineArray[i]))
                        continue;
                    if (lineArray[i] == '"')
                    {
                        flag = true;
                        continue;
                    }
                }
                if (flag && lineArray[i] == '"')
                {
                    if ((i + 1) < line.Length && lineArray[i + 1] == '"')
                        i++;
                    else
                    {
                        if ((i + 1) < line.Length && lineArray[i + 1] != Delimiter)
                            return null;
                        flag = false;
                        continue;
                    }
                }
                if (flag || lineArray[i] != Delimiter)
                    str += lineArray[i];
                else
                {
                    str = str.Trim();
                    list.Add(str);
                    str = string.Empty;
                }
            }
            str = str.Trim();
            list.Add(str);
            return list;
        }
    }
}