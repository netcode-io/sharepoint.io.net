﻿using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharePoint.IO.Managers
{
    /// <summary>
    /// JsInjector
    /// </summary>
    public class JsInjector
    {
        readonly Web _web;
        readonly ClientRuntimeContext _ctx;
        readonly ILogger _log;
        const string DefaultScriptDescription = "Custom Script";
        const string DefaultScriptLocation = "ScriptLink";

        /// <summary>
        /// Initializes a new instance of the <see cref="JsInjector" /> class.
        /// </summary>
        /// <param name="web">The web.</param>
        /// <param name="log">The log.</param>
        public JsInjector(Web web, ILogger log)
        {
            _web = web;
            _ctx = web.Context;
            _log = log;
        }

        /// <summary>
        /// Adds the js link asynchronous.
        /// </summary>
        /// <param name="scripts">The scripts.</param>
        /// <param name="scriptDescription">The script description.</param>
        /// <param name="scriptLocation">The script location.</param>
        /// <param name="allSites">if set to <c>true</c> [all sites].</param>
        public async Task AddJsLinkAsync(IEnumerable<string> scripts, string scriptDescription, string scriptLocation = null, bool allSites = true)
        {
            var b = GenerateJsScriptBlock(_web, scripts);
            await RegisterScriptBlockAsync(_web, b, scriptDescription, scriptLocation ?? DefaultScriptLocation, allSites);
        }

        /// <summary>
        /// Deletes the script links asynchronous.
        /// </summary>
        /// <param name="web">The web.</param>
        /// <param name="scriptLocation">The script location.</param>
        public async Task DeleteScriptLinksAsync(Web web, string scriptLocation = null)
        {
            if (scriptLocation == null)
                scriptLocation = DefaultScriptLocation;
            var actions = web.UserCustomActions.ToArray();
            foreach (var action in actions.Where(action => action.Location == scriptLocation))
            {
                action.DeleteObject();
                await _ctx.ExecuteQueryAsync();
            }
        }

        async Task AddScriptLinkAsync(Web web, StringBuilder b, string scriptDescription, string scriptLocation)
        {
            var action = web.UserCustomActions.Add();
            action.Description = scriptDescription ?? DefaultScriptDescription;
            action.Location = scriptLocation ?? DefaultScriptLocation;
            action.ScriptBlock = b.ToString();
            action.Update();
            _ctx.Load(_web, s => s.UserCustomActions);
            await _ctx.ExecuteQueryAsync();
        }

        async Task RegisterScriptBlockAsync(Web web, StringBuilder b, string scriptDescription, string scriptLocation, bool allSite)
        {
            _ctx.Load(web, s => s.Webs, s => s.UserCustomActions);
            await _ctx.ExecuteQueryAsync();
            await DeleteScriptLinksAsync(web);
            await AddScriptLinkAsync(web, b, scriptDescription, scriptLocation);
            _log?.LogInformation($"JS Injection register for: {web.ServerRelativeUrl}");
            if (allSite)
                foreach (var s in web.Webs)
                    await RegisterScriptBlockAsync(s, b, scriptDescription, scriptLocation, allSite);
            else await RegisterScriptBlockAsync(web.Webs[0], b, scriptDescription, scriptLocation, allSite);
        }

        static StringBuilder GenerateJsScriptBlock(Web web, IEnumerable<string> paths)
        {
            var cacheVersion = Guid.NewGuid().ToString().Replace("-", "");
            var b = new StringBuilder("var headID = document.getElementsByTagName('head')[0]; var newScript;\n");
            foreach (var path in paths)
            {
                var newPath = path.StartsWith("#") ? path.Substring(1) : path;
                newPath = newPath.IndexOf("://") != -1 ? newPath : $"{web.Url.TrimEnd(FileShaman.TrimChars)}/{newPath}";
                if (path.StartsWith("#")) { var id = path.Substring(path.LastIndexOf("/") + 1); b.AppendLine($"RegisterSod('{id}', '{newPath}?v={cacheVersion}');"); }
                else b.AppendLine($"newScript = document.createElement('script'); newScript.type = 'text/javascript'; newScript.src = '{newPath}?v={cacheVersion}'; headID.appendChild(newScript);");
            }
            return b;
        }
    }
}
