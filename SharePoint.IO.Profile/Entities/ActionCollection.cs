using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Xml.Serialization;

namespace SharePoint.IO.Profile.Entities
{
    /// <summary>
    /// Actions Collection
    /// </summary>
    [XmlRoot("actions"), Serializable]
    public class ActionCollection : Collection<BaseAction>
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ActionCollection"/> class.
        /// </summary>
        public ActionCollection() : base() { }
        /// <summary>
        /// Initializes a new instance of the <see cref="ActionCollection"/> class.
        /// </summary>
        /// <param name="list">The list.</param>
        public ActionCollection(IList<BaseAction> list) : base(list) { }
    }
}