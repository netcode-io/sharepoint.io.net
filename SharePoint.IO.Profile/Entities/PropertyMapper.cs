using System;
using System.Xml.Serialization;

namespace SharePoint.IO.Profile.Entities
{
    /// <summary>
    /// The User property mapper instance
    /// </summary>
    [XmlInclude(typeof(PropertyBase))]
    [XmlInclude(typeof(PropertyCollection))]
    [XmlInclude(typeof(BaseAction))]
    [Serializable]
    public class PropertyMapper
    {
        /// <summary>
        /// Gets or sets the action.
        /// </summary>
        /// <value>
        /// The action.
        /// </value>
        public ActionCollection Actions { get; set; }
    }
}