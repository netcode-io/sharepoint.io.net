using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.ObjectModel;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace SharePoint.IO.Profile.Entities
{
    /// <summary>
    /// The abstract base action class
    /// </summary>
    [XmlType("action"), Serializable]
    public abstract class BaseAction
    {
        /// <summary>
        /// Gets or sets the CSV file location.
        /// </summary>
        /// <value>
        /// The CSV file location.
        /// </value>
        public string DirectoryLocation { get; set; } = "Input";

        /// <summary>
        /// Gets or sets the errors.
        /// </summary>
        /// <value>
        /// The errors.
        /// </value>
        public Collection<string> Errors { get; set; }

        /// <summary>
        /// Gets or sets the properties.
        /// </summary>
        /// <value>
        /// The properties.
        /// </value>
        public PropertyCollection Properties { get; set; }

        /// <summary>
        /// Gets or sets the tenant Actions.
        /// </summary>
        /// <value>
        /// The tenant Actions.
        /// </value>
        public ActionCollection Actions { get; set; }

        /// <summary>
        /// Iterates the row from the CSV file
        /// </summary>
        /// <param name="tag">The tag.</param>
        /// <param name="context">The ClientContext instance.</param>
        /// <param name="entries">The collection values per row.</param>
        /// <param name="log">The log.</param>
        public abstract Task IterateCollectionAsync(object tag, ClientContext context, Collection<string> entries, ILogger log);

        /// <summary>
        /// Executes the business logic
        /// </summary>
        /// <param name="log">The log.</param>
        public abstract Task ExecuteAsync(BaseAction parentAction, DateTime CurrentTime, ILogger log);
    }
}