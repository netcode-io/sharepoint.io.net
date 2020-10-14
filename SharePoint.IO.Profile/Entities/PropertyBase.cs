using System;
using System.Xml.Serialization;

namespace SharePoint.IO.Profile.Entities
{
    /// <summary>
    /// The property base class
    /// </summary>
    [XmlType("Property"), Serializable]
    public class PropertyBase
    {
        /// <summary>
        /// Gets or sets the index.
        /// </summary>
        /// <value>
        /// The index.
        /// </value>
        [XmlAttribute("index")] public int Index { get; set; }

        /// <summary>
        /// Gets or sets the name.
        /// </summary>
        /// <value>
        /// The name.
        /// </value>
        [XmlAttribute("name")] public string Name { get; set; }

        /// <summary>
        /// Gets the LDAP mapping.
        /// </summary>
        /// <value>
        /// The mapping.
        /// </value>
        [XmlAttribute("mapping")] public string Mapping { get; set; }

        /// <summary>
        /// Processes the property information
        /// </summary>
        /// <param name="propertyData">The property data.</param>
        /// <param name="value">The value.</param>
        /// <param name="action">The action being executed.</param>
        /// <returns>The parsed property value</returns>
        public virtual object Process(object propertyData, string value, BaseAction action) => value;
    }
}