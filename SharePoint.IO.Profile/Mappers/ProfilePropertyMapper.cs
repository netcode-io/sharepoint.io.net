using SharePoint.IO.Profile.Entities;
using SharePoint.IO.Profile.UserProfileService;

namespace SharePoint.IO.Profile.Mappers
{
    /// <summary>
    /// The Profile Property Mapper
    /// </summary>
    public class ProfilePropertyMapper : PropertyBase
    {
        /// <summary>
        /// Processes the property information
        /// </summary>
        /// <param name="propertyData">The property data.</param>
        /// <param name="value">The value.</param>
        /// <param name="action">The action being executed.</param>
        /// <returns>
        /// The parsed property value
        /// </returns>
        public override object Process(object propertyData, string value, BaseAction action)
        {
            if (propertyData is PropertyData data)
            {
                data.IsValueChanged = true;
                data.Values = new[] { new ValueData { Value = value } };
                return data;
            }
            return value;
        }
    }
}