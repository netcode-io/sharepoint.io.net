using Contoso.Extensions.Services;
using Dapper;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using SharePoint.IO.Profile.Entities;
using SharePoint.IO.Profile.Services;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Linq;

namespace SharePoint.IO.Profile.Mappers
{
    /// <summary>
    /// SQL Connector
    /// </summary>
    public class SqlConnector : BaseAction
    {
        static readonly IDbService _db = new DbService();
        const int ACCOUNTDISABLE = 0x0002;
        int _totalUsers = 0;
        int _totalFailures = 0;

        public string SPOAccountUPN { get; set; }
        public IConfiguration Configuration { get; set; }
        public string ConnectionName { get; set; } = "Main";
        public string SPOClaimsString { get; set; } = "i:0#.f|membership|";
        public int PageSize { get; set; } = 100;
        public string StoredProcedure { get; set; }
        public int CommandTimeout { get; set; } = 60;
        public int UserNameIndex { get; set; }

        /// <summary>
        /// Iterates the row from the CSV file
        /// </summary>
        /// <param name="tag">The tag.</param>
        /// <param name="context">The ClientContext instance.</param>
        /// <param name="entries">The collection values per row.</param>
        /// <param name="log">The log.</param>
        /// <exception cref="NotImplementedException"></exception>
        /// <exception cref="System.NotImplementedException"></exception>
        public override void IterateCollection(object tag, Microsoft.SharePoint.Client.ClientContext context, Collection<string> entries, ILogger log) => throw new NotImplementedException();

        /// <summary>
        /// Executes the SQL logic
        /// </summary>
        /// <param name="parentAction">Inherit parent properties = null</param>
        /// <param name="currentTime">Locked program timestamp value</param>
        /// <param name="log">The log.</param>
        public override void Execute(BaseAction parentAction, DateTime currentTime, ILogger log)
        {
            ExtractSqlResults(log, currentTime);
            log.LogInformation($"Successfully extracted {_totalUsers} user objects from {ConnectionName} with {_totalFailures} failures");
        }

        /// <summary>
        /// Tries the parse value.
        /// </summary>
        /// <param name="log">The log.</param>
        /// <param name="item">The item.</param>
        /// <param name="row">The entry.</param>
        /// <param name="value">The value.</param>
        /// <param name="accountDisabled">if set to <c>true</c> [account disabled].</param>
        /// <param name="attr">The attribute.</param>
        /// <returns></returns>
        protected virtual bool TryParseValue(ILogger log, PropertyBase item, CsvRow row, string value, IDictionary<string, object> sre) => false;

        /// <summary>
        /// Performs SQL query and extracts attributes.
        /// </summary>
        /// <param name="log">The log.</param>
        /// <param name="currentTime">Locked program timestamp value</param>
        private void ExtractSqlResults(ILogger log, DateTime currentTime)
        {
            log.LogInformation($"Establishing SQL Connection to: {ConnectionName}");
            using (var connection = CreateSqlConnection())
            {
                log.LogInformation($"Performing a sql operation on: {StoredProcedure}");
                List<dynamic> set = null;
                try { set = connection.Query(StoredProcedure, null, commandType: CommandType.StoredProcedure, commandTimeout: CommandTimeout).ToList(); }
                catch (Exception e) { throw new Exception("An error occurred whilst querying", e); }

                foreach (var items in GroupAt(set, PageSize, x => x))
                {
                    var responseCount = items.Count();
                    var currentBatchSize = responseCount != PageSize ? responseCount : PageSize;
                    var filePath = CsvCreateFile(DirectoryLocation, _totalUsers, currentBatchSize);

                    // Create CSV file for current batch of users
                    using (var batchFile = new CsvWriter(filePath))
                    {
                        // Create column headings for CSV file
                        var heading = new CsvRow();
                        // Iterate over attribute headings
                        foreach (var item in Properties)
                            heading.Add(item.Name);
                        batchFile.CsvWrite(heading, log);
                        // Create new CSV row for each user
                        foreach (IDictionary<string, object> sre in items)
                        {
                            // Placeholder for CSV entry of current user
                            var entry = new CsvRow();
                            // Extract each user attribute specified in XML file
                            foreach (var item in Properties)
                            {
                                try
                                {
                                    var value = sre[item.Mapping] != null ? sre[item.Mapping].ToString() : string.Empty;
                                    if (TryParseValue(log, item, entry, value, sre))
                                        continue;
                                    if (item.Index == UserNameIndex)
                                    {
                                        entry.Add(CreateUserAccountName(value));
                                        continue;
                                    }
                                    entry.Add(value);
                                }
                                catch (Exception e)
                                {
                                    if (log != null)
                                    {
                                        log.LogCritical(e, string.Empty);
                                        _totalFailures++;
                                    }
                                }
                            }
                            // Write current user to CSV file
                            batchFile.CsvWrite(entry, log);
                            // Increment user count value
                            _totalUsers++;
                        }
                        log.LogInformation($"Successfully extracted {currentBatchSize} users to {filePath} - the total user count is: {_totalUsers}");
                    }
                }
            }
        }

        string CreateUserAccountName(string value)
        {
            var position = value.IndexOf('\\');
            if (position > 0)
                value = value.Substring(position + 1);
            return $"{SPOClaimsString}{value}@{SPOAccountUPN}";
        }

        /// <summary>
        /// Create the CSV batch file.
        /// </summary>
        /// <param name="location">Directory location.</param>
        /// <param name="totalCount">Total number of users.</param>
        public string CsvCreateFile(string location, int totalCount, int currentBatchSize)
        {
            if (!Directory.Exists(location))
                Directory.CreateDirectory(location);
            var startValue = (totalCount == 0 ? 1 : totalCount);
            return $"{location}\\SqlConnector-{startValue}-{totalCount + currentBatchSize}.csv";
        }

        /// <summary>
        /// Establish the SQL Connection
        /// </summary>
        public IDbConnection CreateSqlConnection() => _db.GetConnection(ConnectionName);

        static IEnumerable<TResult> GroupAt<TSource, TResult>(IEnumerable<TSource> source, int size, Func<IEnumerable<TSource>, TResult> resultSelector)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));
            if (size <= 0)
                throw new ArgumentOutOfRangeException(nameof(size));
            if (resultSelector == null)
                throw new ArgumentNullException(nameof(resultSelector));
            var index = 0;
            var items = new TSource[size];
            foreach (var item in source)
            {
                items[index++] = item;
                if (index != size)
                    continue;
                yield return resultSelector(items);
                index = 0;
                items = new TSource[size];
            }
            if (index > 0)
                yield return resultSelector(items.Take(index));
        }
    }
}