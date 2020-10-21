using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SharePoint.IO.Profile
{
    /// <summary>
    /// Extensions
    /// </summary>
    public static class Extensions
    {
        internal static IEnumerable<TResult> GroupAt<TSource, TResult>(this IEnumerable<TSource> source, int size, Func<IEnumerable<TSource>, TResult> resultSelector)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));
            if (size <= 0)
                throw new ArgumentOutOfRangeException(nameof(size));
            if (resultSelector == null)
                throw new ArgumentNullException(nameof(resultSelector));
            var index = 0;
            var items = new TSource[size];
            foreach (var item in source)
            {
                items[index++] = item;
                if (index != size)
                    continue;
                yield return resultSelector(items);
                index = 0;
                items = new TSource[size];
            }
            if (index > 0)
                yield return resultSelector(items.Take(index));
        }

        /// <summary>
        /// Parses the month date.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns></returns>
        /// <exception cref="InvalidDataException">Unable to parse MonthDate.  Inner Exception: {e.Message}</exception>
        public static string ParseMonthDate(this string value)
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

        /// <summary>
        /// Parses the date no year.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns></returns>
        /// <exception cref="InvalidDataException">Unable to parse DateNoYear.</exception>
        public static string ParseDateNoYear(this string value) => string.IsNullOrEmpty(value)
            ? string.Empty
            : DateTime.TryParse(value, out var valueDate)
            ? new DateTime(2000, valueDate.Month, valueDate.Day).ToString()
            : throw new InvalidDataException("Unable to parse DateNoYear.");

        /// <summary>
        /// Parses the postal address.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <returns></returns>
        public static string ParsePostalAddress(this string value)
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