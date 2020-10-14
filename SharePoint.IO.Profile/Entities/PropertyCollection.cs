using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Xml.Serialization;

namespace SharePoint.IO.Profile.Entities
{
    /// <summary>
    /// The collection of properties
    /// </summary>
    [XmlRoot("properties"), Serializable]
    public class PropertyCollection : Collection<PropertyBase>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="PropertyCollection"/> class.
        /// </summary>
        public PropertyCollection() : base() { }
        /// <summary>
        /// Initializes a new instance of the <see cref="PropertyCollection"/> class.
        /// </summary>
        /// <param name="list">The collection of properties.</param>
        public PropertyCollection(IList<PropertyBase> list) : base(list) { }
    }
}