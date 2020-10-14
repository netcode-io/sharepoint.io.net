//using Microsoft.Extensions.Logging;
//using Microsoft.SharePoint.Client;
//using SharePoint.IO.Profile.Entities;
//using System;
//using System.Collections.Generic;
//using System.Collections.ObjectModel;
//using System.IO;
//using System.Net;
//using System.Security;
//using System.Threading;

//namespace SharePoint.IO.Profile.Mappers
//{
//    /// <summary>
//    /// The user profile mapper instance
//    /// </summary>
//    public class UserProfileMapper : BaseAction
//    {
//        /// <summary>
//        /// The fed authentication cookie name
//        /// </summary>
//        const string FedAuthCookieName = "SPOIDCRL";
//        //const string FedAuthCookieName = "FedAuth";
//        const string AuthenticationService = "/_vti_bin/authentication.asmx";
//        /// <summary>
//        /// The profile service
//        /// </summary>
//        const string ProfileService = "/_vti_bin/userprofileservice.asmx";
//        /// <summary>
//        /// The SharePoint Online cookie value
//        /// </summary>
//        const string SPOIDCookieValue = "SPOIDCRL=";
//        //const string SPOIDCookieValue = "FedAuth=";

//        /// <summary>
//        /// Gets or sets the tenant site URL.
//        /// </summary>
//        /// <value>
//        /// The tenant site URL.
//        /// </value>
//        public string TenantSiteUrl { get; set; }

//        /// <summary>
//        /// Gets or sets the tenant admin login.
//        /// </summary>
//        /// <value>
//        /// The tenant admin login.
//        /// </value>
//        public string TenantAdminLogin { get; set; }

//        /// <summary>
//        /// Gets or sets the tenant admin password.
//        /// </summary>
//        /// <value>
//        /// The tenant admin password.
//        /// </value>
//        public string TenantAdminPassword { get; set; }

//        /// <summary>
//        /// Gets or sets the name of the tenant admin user.
//        /// </summary>
//        /// <value>
//        /// The name of the tenant admin user.
//        /// </value>
//        public string TenantAdminUserName { get; set; }

//        /// <summary>
//        /// Gets or sets the index of the user name.
//        /// </summary>
//        /// <value>
//        /// The index of the user name.
//        /// </value>
//        public int UserNameIndex { get; set; }
//        public int SleepPeriod { get; set; }
//        public bool Keep { get; set; }

//        /// <summary>
//        /// Iterates the row from the CSV file
//        /// </summary>
//        /// <param name="context">The ClientContext instance.</param>
//        /// <param name="entries">The collection values per row.</param>
//        /// <param name="log">The log.</param>
//        public override void IterateCollection(object tag, ClientContext context, Collection<string> entries, ILogger log)
//        {
//            var profileService = (UserProfileService.UserProfileService)tag;
//            var data = new List<PropertyData>();
//            foreach (var item in Properties)
//                if (item.Index < entries.Count)
//                {
//                    try
//                    {
//                        var account = entries[UserNameIndex];
//                        var property = new PropertyData { Name = item.Name };
//                        property = item.Process(property, entries[item.Index], this) as PropertyData;
//                        data.Add(property);
//                    }
//                    catch (Exception e) { log.LogCritical(e, $"Error occured whilst processing account '{entries[UserNameIndex]}', Property '{item.Name}'. Stack {e.StackTrace}"); }
//                }
//            log.LogInformation($"Attempting to update profile for account '{entries[UserNameIndex]}'");
//            try
//            {
//                profileService.ModifyUserPropertyByAccountName(entries[UserNameIndex], data.ToArray());
//                log.LogInformation(entries[UserNameIndex], "SUCCESS");
//            }
//            catch (Exception e)
//            {
//                log.LogCritical(e, $"Error occured whilst processing account '{entries[UserNameIndex]}' - the account does not exist. InnerException: {e.Message}");
//                log.LogInformation(entries[UserNameIndex], "FAILURE");
//            }
//        }

//        /// <summary>
//        /// Executes the business logic
//        /// </summary>
//        /// <param name="log">The log.</param>
//        public override void Execute(BaseAction parentAction, DateTime currentTime, ILogger log)
//        {
//            var credential = !string.IsNullOrEmpty(TenantAdminLogin) ? ConfigurationManagerEx.Decode<NetworkCredential, string>(TenantAdminLogin, null) : new NetworkCredential(TenantAdminUserName, TenantAdminPassword);
//            if (parentAction != null)
//                Properties = parentAction.Properties;
//            var csvProcessor = new CsvProcessor();
//            var csvFiles = Directory.GetFiles(DirectoryLocation, "*.csv", SearchOption.TopDirectoryOnly);
//            log.LogInformation($"Attempting to get files from directory 'location' {DirectoryLocation}. Number of files found {csvFiles.Length}");
//            foreach (string csvFile in csvFiles)
//            {
//                log.LogInformation($"Attempting to read CSV file '{csvFile0}' from location {DirectoryLocation}");
//                log.LogInformation($"Pausing the utility for '{SleepPeriod}' seconds so ASMX service is not overloaded");
//                Thread.Sleep(SleepPeriod * 1000);
//                using (var reader = new StreamReader(csvFile))
//                {
//                    log.LogInformation($"Establishing connection with tenant at '{TenantSiteUrl}'");
//                    using (var context = new ClientContext(TenantSiteUrl))
//                    {
//                        var site = new Uri(TenantSiteUrl);
//                        try
//                        {
//                            log.LogInformation(site.ToString() + ProfileService);
//                            using (var profileService = new UserProfileService.UserProfileService { Url = site.ToString() + ProfileService })
//                            {
//                                profileService.UseDefaultCredentials = false;
//                                using (var password = new SecureString())
//                                {
//                                    foreach (char c in credential.Password.ToCharArray())
//                                        password.AppendChar(c);
//                                    log.LogInformation($"Attempting to authenticate against tenant with user name '{credential.UserName}'");
//                                    var credentials = new SharePointOnlineCredentials(credential.UserName, password);
//                                    var cookie = credentials.GetAuthenticationCookie(site);
//                                    profileService.CookieContainer = new CookieContainer();
//                                    profileService.CookieContainer.Add(new Cookie(FedAuthCookieName, cookie.TrimStart(SPOIDCookieValue.ToCharArray()), string.Empty, site.Authority));
//                                    csvProcessor.Execute(reader, (entries, y) => { IterateCollection(profileService, context, entries, logger); }, log);
//                                }
//                            }
//                        }
//                        catch (Exception e) { log.LogCritical(e, e.Message); }
//                    }
//                }
//                if (!Keep)
//                    System.IO.File.Delete(csvFile); // Clean up current CSV file
//            }
//        }
//    }
//}