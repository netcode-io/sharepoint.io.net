using Microsoft.Extensions.Logging;
using SharePoint.IO.Profile.Entities;
using SharePoint.IO.Profile.Services;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.DirectoryServices.Protocols;
using System.IO;
using System.Linq;
using System.Net;
using System.Security;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;

namespace SharePoint.IO.Profile.Mappers
{
    /// <summary>
    /// LDAP Connector
    /// </summary>
    public class LdapConnector : BaseAction
    {
        const int ACCOUNTDISABLE = 0x0002;
        int _totalUsers = 0;
        int _totalFailures = 0;

        public string SPOAccountUPN { get; set; }
        public string ServerName { get; set; }
        public string PortNumber { get; set; } = "389";
        public string ServiceLogin { get; set; }
        public string ServiceUserName { get; set; }
        public string ServicePassword { get; set; }
        public string SearchRoot { get; set; }
        public string BatchAction { get; set; } = "bulk";
        public string SPOClaimsString { get; set; } = "i:0#.f|membership|";
        public int PageSize { get; set; } = 100;
        public int QueryTimeout { get; set; } = 60;
        public string LDAPAuthType { get; set; } = "Basic";
        public string DirectoryType { get; set; } = "DirectoryServer";
        public int ProtocolVers { get; set; } = 3;
        public string CertificatePath { get; set; }
        public int DeltaPeriod { get; set; } = 2;
        public int UserNameIndex { get; set; }

        /// <summary>
        /// Iterates the row from the CSV file
        /// </summary>
        /// <param name="tag">The tag.</param>
        /// <param name="context">The ClientContext instance.</param>
        /// <param name="entries">The collection values per row.</param>
        /// <exception cref="System.NotImplementedException"></exception>
        public override Task IterateCollectionAsync(object tag, Microsoft.SharePoint.Client.ClientContext context, Collection<string> entries) => throw new NotImplementedException();

        /// <summary>
        /// Executes the LDAP logic
        /// </summary>
        /// <param name="parentAction">Inherit parent properties = null</param>
        /// <param name="currentTime">Locked program timestamp value</param>
        public override async Task ExecuteAsync(BaseAction parentAction, DateTime currentTime)
        {
            await ExtractLdapResultsAsync(currentTime);
            Log?.LogInformation($"Successfully extracted {_totalUsers} user objects from {ServerName} with {_totalFailures} failures");
        }

        /// <summary>
        /// Tries the parse value.
        /// </summary>
        /// <param name="logger">The logger.</param>
        /// <param name="item">The item.</param>
        /// <param name="entry">The entry.</param>
        /// <param name="value">The value.</param>
        /// <param name="accountDisabled">if set to <c>true</c> [account disabled].</param>
        /// <param name="attr">The attribute.</param>
        /// <returns></returns>
        protected virtual bool TryParseValue(PropertyBase item, CsvRow entry, string value, bool accountDisabled, DirectoryAttribute attr, SearchResultEntry sre) => false;

        /// <summary>
        /// Performs LDAP Search and extracts attributes.
        /// </summary>
        /// <param name="logger">The logger.</param>
        /// <param name="currentTime">Locked program timestamp value</param>
        Task ExtractLdapResultsAsync(DateTime currentTime)
        {
            var attributesToAdd = new List<string>();
            foreach (var item in Properties)
                attributesToAdd.Add(item.Mapping);
            attributesToAdd.Add("userAccountControl");
            var ldapFilter = SetQueryFilter(BatchAction, currentTime);
            var searchRequest = new SearchRequest(SearchRoot, ldapFilter, SearchScope.Subtree, attributesToAdd.ToArray());
            var pageResponse = new PageResultRequestControl(PageSize);
            var searchOptions = new SearchOptionsControl(System.DirectoryServices.Protocols.SearchOption.DomainScope);
            searchRequest.Controls.Add(pageResponse);
            searchRequest.Controls.Add(searchOptions);
            Log?.LogInformation($"Establishing LDAP Connection to: {ServerName}");
            using (var connection = CreateLdapConnection())
            {
                Log?.LogInformation($"Performing a {BatchAction} operation with filter: {ldapFilter}");
                while (true)
                {
                    SearchResponse response = null;
                    try { response = connection.SendRequest(searchRequest) as SearchResponse; }
                    catch (Exception e) { throw new Exception("An error occurred whilst creating the SearchResponse", e); }

                    var responseCount = response.Entries.Count;
                    var currentBatchSize = (responseCount != PageSize ? responseCount : PageSize);
                    var filePath = CsvCreateFile(DirectoryLocation, _totalUsers, currentBatchSize);

                    foreach (var control in response.Controls)
                        if (control is PageResultResponseControl responseControl)
                        {
                            pageResponse.Cookie = responseControl.Cookie;
                            break;
                        }

                    // Create CSV file for current batch of users
                    using (var batchFile = new CsvWriter(filePath, Log))
                    {
                        // Create column headings for CSV file
                        var heading = new CsvRow();
                        // Iterate over attribute headings
                        foreach (var item in Properties)
                            heading.Add(item.Name);
                        batchFile.CsvWrite(heading);
                        // Create new CSV row for each user
                        foreach (SearchResultEntry sre in response.Entries)
                        {
                            // Placeholder for CSV entry of current user
                            var entry = new CsvRow();
                            var syncValue = true;

                            // Get whether account is disabled
                            var userAccountControlAttr = sre.Attributes["userAccountControl"];
                            var accountDisabled = true;
                            if (userAccountControlAttr != null && userAccountControlAttr.Count > 0)
                            {
                                var userAccountControlValue = userAccountControlAttr[0].ToString();
                                try
                                {
                                    var userAccountControl = int.Parse(userAccountControlValue);
                                    accountDisabled = (userAccountControl & ACCOUNTDISABLE) == ACCOUNTDISABLE;
                                }
                                catch (Exception e) { Log?.LogCritical(e, e.Message); }
                            }

                            // Extract each user attribute specified in XML file
                            foreach (var item in Properties)
                                try
                                {
                                    var attr = sre.Attributes[item.Mapping];
                                    var value = (attr != null && attr.Count > 0 ? attr[0].ToString() : string.Empty);
                                    if (syncValue && TryParseValue(item, entry, value, accountDisabled, attr, sre))
                                        continue;
                                    if (item.Index == UserNameIndex)
                                    {
                                        entry.Add(CreateUserAccountName(value, attr));
                                        continue;
                                    }
                                    entry.Add(syncValue ? value : string.Empty);
                                }
                                catch (Exception e) { Log?.LogCritical(e, string.Empty); _totalFailures++; }
                            // Write current user to CSV file
                            batchFile.CsvWrite(entry);
                            // Increment user count value
                            _totalUsers++;
                        }
                    }
                    Log?.LogInformation($"Successfully extracted {currentBatchSize} users to {filePath} - the total user count is: {_totalUsers}");
                    if (pageResponse.Cookie.Length == 0)
                        break;
                }
            }
            return Task.CompletedTask;
        }

        string CreateUserAccountName(string value, DirectoryAttribute attr)
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
            var startValue = totalCount == 0 ? 1 : totalCount;
            return $"{location}\\LdapConnector-{startValue}-{totalCount + currentBatchSize}.csv";
        }

        /// <summary>
        /// Establish the LDAP Connection
        /// </summary>
        public LdapConnection CreateLdapConnection()
        {
            var identifier = CreateIdentifier();
            var credential = !string.IsNullOrEmpty(ServiceLogin)
                ? CredentialManager.ReadGeneric(ServiceLogin)
                : new NetworkCredential(ServiceUserName, ServicePassword);
            var authType = (AuthType)Enum.Parse(typeof(AuthType), LDAPAuthType);
            var connection = new LdapConnection(identifier, credential, authType);
            connection.SessionOptions.ProtocolVersion = ProtocolVers;
            /*
            connection.SessionOptions.VerifyServerCertificate = new VerifyServerCertificateCallback(VerifyServerCertificate);
            connection.SessionOptions.QueryClientCertificate = new QueryClientCertificateCallback(QueryClientCertificate);
            var timeSpan = new TimeSpan(0, 0, QueryTimeout);
            connection.Timeout = timeSpan;
            connection.SessionOptions.SecureSocketLayer = true;
            */
            connection.Bind();
            return connection;
        }

        /// <summary>
        /// Sets the query filter for LDAP search
        /// </summary>
        /// <param name="action">Flag to toggle search query scoper</param>
        /// <param name="currentTime">Locked time stamp when program started</param>
        public string SetQueryFilter(string action, DateTime currentTime)
        {
            string query;
            // Apply filter scope based on IsBulk value
            if (action == "bulk")
                query = "(objectCategory=Person)"; // Extract all users and attributes 
            else if (action == "delta")
            {
                var timeNow = currentTime.ToString("yyyyMMddHHmmss.0Z");
                var timeDelta = currentTime.AddDays(-DeltaPeriod).ToString("yyyyMMddHHmmss.0Z");
                query = "(&(whenChanged>=" + timeDelta + ")(whenChanged<=" + timeNow + ")(objectCategory=Person))";
            }
            else
                query = action;
            return query;
        }

        LdapDirectoryIdentifier CreateIdentifier()
        {
            var hostEntry = Dns.GetHostEntry(ServerName);
            if (hostEntry == null)
                throw new InvalidOperationException($"Unable to find host {ServerName}");
            var address = hostEntry.AddressList.FirstOrDefault();
            var connectionString = address.ToString() + ":" + PortNumber;
            return new LdapDirectoryIdentifier(connectionString);
        }

        bool VerifyServerCertificate(LdapConnection connection, X509Certificate certificate)
        {
            //var id = connection.Directory as LdapDirectoryIdentifier;
            var store = new X509Store("CertStore", StoreLocation.LocalMachine);
            store.Open(OpenFlags.ReadWrite);
            var newcert = new X509Certificate2(certificate);
            store.Add(newcert);
            return true;
        }

        X509Certificate QueryClientCertificate(LdapConnection connection, byte[][] trustedCAs)
        {
            //var id = connection.Directory as LdapDirectoryIdentifier;
            if (IsTrustedContosoCA(trustedCAs))
            {
                var cert = new X509Certificate();
                cert.Import(GetPath(CertificatePath), ServicePassword, X509KeyStorageFlags.DefaultKeySet);
                connection.ClientCertificates.Add(cert);
                return null;
            }
            return null;
        }

        bool IsTrustedContosoCA(byte[][] trustedCAs)
        {
            foreach (var ca in trustedCAs)
            {
                var utf8 = System.Text.Encoding.UTF8.GetString(ca);
                if (utf8.ToLower().Contains("contoso"))
                    return true;
            }
            return false;
        }

        string GetPath(string path)
        {
            if (path.StartsWith("~/"))
                path = Path.Combine(Environment.CurrentDirectory, path.Substring(2));
            else if (!Path.IsPathRooted(path))
                path = Path.Combine(Environment.CurrentDirectory, path);
            return path;
        }
    }
}