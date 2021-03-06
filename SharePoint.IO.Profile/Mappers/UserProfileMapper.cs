﻿using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using SharePoint.IO.Profile.Entities;
using SharePoint.IO.Profile.Services;
using SharePoint.IO.Profile.UserProfileService;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Net;
using System.Security;
using System.Threading;
using System.Threading.Tasks;

namespace SharePoint.IO.Profile.Mappers
{
    /// <summary>
    /// The user profile mapper instance
    /// </summary>
    public class UserProfileMapper : BaseAction
    {
        /// <summary>
        /// The fed authentication cookie name
        /// </summary>
        const string FedAuthCookieName = "SPOIDCRL";
        //const string FedAuthCookieName = "FedAuth";
        const string AuthenticationService = "/_vti_bin/authentication.asmx";
        /// <summary>
        /// The profile service
        /// </summary>
        const string ProfileService = "/_vti_bin/userprofileservice.asmx";
        /// <summary>
        /// The SharePoint Online cookie value
        /// </summary>
        const string SPOIDCookieValue = "SPOIDCRL=";
        //const string SPOIDCookieValue = "FedAuth=";

        /// <summary>
        /// Gets or sets the tenant site URL.
        /// </summary>
        /// <value>
        /// The tenant site URL.
        /// </value>
        public string TenantSiteUrl { get; set; }

        /// <summary>
        /// Gets or sets the tenant admin login.
        /// </summary>
        /// <value>
        /// The tenant admin login.
        /// </value>
        public string TenantAdminLogin { get; set; }

        /// <summary>
        /// Gets or sets the tenant admin password.
        /// </summary>
        /// <value>
        /// The tenant admin password.
        /// </value>
        public string TenantAdminPassword { get; set; }

        /// <summary>
        /// Gets or sets the name of the tenant admin user.
        /// </summary>
        /// <value>
        /// The name of the tenant admin user.
        /// </value>
        public string TenantAdminUserName { get; set; }

        /// <summary>
        /// Gets or sets the index of the user name.
        /// </summary>
        /// <value>
        /// The index of the user name.
        /// </value>
        public int UserNameIndex { get; set; }
        /// <summary>
        /// Gets or sets the sleep period.
        /// </summary>
        /// <value>
        /// The sleep period.
        /// </value>
        public int SleepPeriod { get; set; }
        /// <summary>
        /// Gets or sets a value indicating whether this <see cref="UserProfileMapper"/> is keep.
        /// </summary>
        /// <value>
        ///   <c>true</c> if keep; otherwise, <c>false</c>.
        /// </value>
        public bool Keep { get; set; }

        /// <summary>
        /// Iterates the row from the CSV file
        /// </summary>
        /// <param name="tag">The tag.</param>
        /// <param name="context">The ClientContext instance.</param>
        /// <param name="entries">The collection values per row.</param>
        public override async Task IterateCollectionAsync(object tag, ClientContext context, Collection<string> entries)
        {
            var profileService = (UserProfileServiceSoapClient)tag;
            var data = new List<PropertyData>();
            foreach (var item in Properties)
                if (item.Index < entries.Count)
                {
                    try
                    {
                        var account = entries[UserNameIndex];
                        var property = new PropertyData { Name = item.Name };
                        property = item.Process(property, entries[item.Index], this) as PropertyData;
                        data.Add(property);
                    }
                    catch (Exception e) { Log?.LogCritical(e, $"Error occured whilst processing account '{entries[UserNameIndex]}', Property '{item.Name}'. Stack {e.StackTrace}"); }
                }
            Log?.LogInformation($"Attempting to update profile for account '{entries[UserNameIndex]}'");
            try
            {
                await profileService.ModifyUserPropertyByAccountNameAsync(entries[UserNameIndex], data.ToArray());
                Log?.LogInformation(entries[UserNameIndex], "SUCCESS");
            }
            catch (Exception e)
            {
                Log?.LogCritical(e, $"Error occured whilst processing account '{entries[UserNameIndex]}' - the account does not exist. InnerException: {e.Message}");
                Log?.LogInformation(entries[UserNameIndex], "FAILURE");
            }
        }

        /// <summary>
        /// Executes the business logic
        /// </summary>
        /// <param name="log">The log.</param>
        public override Task ExecuteAsync(BaseAction parentAction, DateTime currentTime)
        {
            var credential = !string.IsNullOrEmpty(TenantAdminLogin)
                ? CredentialManager.ReadGeneric(TenantAdminLogin)
                : new NetworkCredential(TenantAdminUserName, TenantAdminPassword);
            if (parentAction != null)
                Properties = parentAction.Properties;
            var csvReader = new CsvReader(Log);
            var csvFiles = Directory.GetFiles(DirectoryLocation, "*.csv", SearchOption.TopDirectoryOnly);
            Log?.LogInformation($"Attempting to get files from directory 'location' {DirectoryLocation}. Number of files found {csvFiles.Length}");
            foreach (string csvFile in csvFiles)
            {
                Log?.LogInformation($"Attempting to read CSV file '{csvFile}' from location {DirectoryLocation}");
                Log?.LogInformation($"Pausing the utility for '{SleepPeriod}' seconds so ASMX service is not overloaded");
                Thread.Sleep(SleepPeriod * 1000);
                using (var reader = new StreamReader(csvFile))
                {
                    Log?.LogInformation($"Establishing connection with tenant at '{TenantSiteUrl}'");
                    using (var context = new ClientContext(TenantSiteUrl))
                    {
                        var site = new Uri(TenantSiteUrl);
                        try
                        {
                            Log?.LogInformation($"{site}{ProfileService}");
                            var profileService = new UserProfileServiceSoapClient(UserProfileServiceSoapClient.EndpointConfiguration.UserProfileServiceSoap12, $"{site}{ProfileService}");
                            {
                                //profileService.UseDefaultCredentials = false;
                                //using (var password = new SecureString())
                                //{
                                //    foreach (char c in credential.Password.ToCharArray())
                                //        password.AppendChar(c);
                                //    log.LogInformation($"Attempting to authenticate against tenant with user name '{credential.UserName}'");
                                //    var credentials = new SharePointOnlineCredentials(credential.UserName, password);
                                //    var cookie = credentials.GetAuthenticationCookie(site);
                                //    profileService.CookieContainer = new CookieContainer();
                                //    profileService.CookieContainer.Add(new Cookie(FedAuthCookieName, cookie.TrimStart(SPOIDCookieValue.ToCharArray()), string.Empty, site.Authority));
                                csvReader.Execute(reader, async (entries, y) => { await IterateCollectionAsync(profileService, context, entries); });
                                //}
                            }
                        }
                        catch (Exception e) { Log?.LogCritical(e, e.Message); }
                    }
                }
                if (!Keep)
                    System.IO.File.Delete(csvFile); // Clean up current CSV file
            }
            return Task.CompletedTask;
        }
    }
}