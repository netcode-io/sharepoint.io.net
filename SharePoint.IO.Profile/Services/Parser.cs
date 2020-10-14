using System;
using System.IO;

namespace SharePoint.IO.Profile.Services
{
    /// <summary>
    /// Parser
    /// </summary>
    public class Parser
    {
        public static string ParseMonthDate(string value)
        {
            if (string.IsNullOrEmpty(value))
                return string.Empty;
            var date = new DateTime().AddYears(1999);
            try
            {
                var monthValue = value.Substring(0, 2);
                var dayValue = value.Substring(2, 2);
                if (int.TryParse(monthValue, out var month) && int.TryParse(dayValue, out var day))
                {
                    date = date.AddMonths(month - 1);
                    date = date.AddDays(day - 1);
                }
            }
            catch (Exception e) { throw new InvalidDataException($"Unable to parse MonthDate.  Inner Exception: {e.Message}"); }
            return date.ToString();
        }

        public static string ParseDateNoYear(string value)
        {
            if (string.IsNullOrEmpty(value))
                return string.Empty;
            if (DateTime.TryParse(value, out var valueDate))
            {
                var date = new DateTime(2000, valueDate.Month, valueDate.Day);
                return date.ToString();
            }
            throw new InvalidDataException("Unable to parse DateNoYear.");
        }

        public static string ParsePostalAddress(string value)
        {
            if (string.IsNullOrEmpty(value))
                return string.Empty;
            var splitValue = value.Split('|');
            if (splitValue.Length == 4) return string.Format("{0}, {1}, {2} {3}", splitValue);
            if (splitValue.Length == 5) return string.Format("{0}, {1}, {2}, {3} {4}", splitValue);
            return string.Empty;
        }
    }
}