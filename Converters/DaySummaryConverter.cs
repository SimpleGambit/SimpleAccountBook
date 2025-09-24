using SimpleAccountBook.Models;
using SimpleAccountBook.ViewModels;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Windows.Data;

namespace SimpleAccountBook.Converters
{
    public class DaySummaryConverter : IMultiValueConverter
    {
        public object? Convert(object[] values, Type targetType, object? parameter, CultureInfo culture)
        {
            Debug.WriteLine(
               "[DaySummaryConverter.Convert] *** CALLED *** with " + values.Length + " values");

            for (int i = 0; i < values.Length && i < 4; i++)
            {
                var val = values[i];
                Debug.WriteLine($"  Value[{i}]: {val?.GetType().Name ?? "null"} = {val}");
            }

            if (values.Length < 2)
            {
                Debug.WriteLine("[DaySummaryConverter] Not enough values, returning Blank");
                return DailySummary.Blank;
            }

            var displayMonth = FindDisplayMonth(values);
            Debug.WriteLine($"[DaySummaryConverter] DisplayMonth = {displayMonth}");

            if (!TryResolveDate(values[0], displayMonth, out var date))
            {
                Debug.WriteLine("[DaySummaryConverter] Failed to resolve date");
                return DailySummary.Blank;
            }
            Debug.WriteLine($"[DaySummaryConverter] Resolved date = {date:yyyy-MM-dd}");

            IReadOnlyDictionary<DateOnly, DailySummary>? dict = values[1] switch
            {
                IReadOnlyDictionary<DateOnly, DailySummary> d => d,
                AccountBookViewModel vm => vm.DailySummaries,
                _ => null
            };

            if (dict is null)
            {
                Debug.WriteLine("[DaySummaryConverter] Dictionary is NULL");
                return DailySummary.Blank;
            }
            Debug.WriteLine($"[DaySummaryConverter] Dictionary contains {dict.Count} entries");

            var key = DateOnly.FromDateTime(date);
            if (dict.TryGetValue(key, out var hit))
            {
                Debug.WriteLine($"[DaySummaryConverter] *** HIT *** for {key}:");
                Debug.WriteLine($"  Income={hit.IncomeAmount:#,0}, Expense={hit.ExpenseAmount:#,0}");
                Debug.WriteLine($"  IncomeText='{hit.IncomeText}', ExpenseText='{hit.ExpenseText}'");
                return hit;
            }

            Debug.WriteLine($"[DaySummaryConverter] No data for {key}");

            if (displayMonth.HasValue && date.Year == displayMonth.Value.Year && date.Month == displayMonth.Value.Month)
            {
                Debug.WriteLine($"[DaySummaryConverter] Creating empty for {date:yyyy-MM-dd}");
                return DailySummary.CreateEmpty(date);
            }

            Debug.WriteLine("[DaySummaryConverter] Returning Blank");
            return DailySummary.Blank;
        }

        public object[] ConvertBack(object? value, Type[] targetTypes, object? parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }

        private static bool TryResolveDate(object? value, DateTime? displayMonth, out DateTime date)
        {
            if (TryExtractDate(value, out date))
            {
                return true;
            }

            if (displayMonth is DateTime anchor)
            {
                if (value is int numericDay && TryCreateDate(anchor, numericDay, out date))
                {
                    return true;
                }

                if (value is string text)
                {
                    if (int.TryParse(text, NumberStyles.Integer, CultureInfo.CurrentCulture, out var parsedDay) &&
     TryCreateDate(anchor, parsedDay, out date))
                    {
                        return true;
                    }

                    if (int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out parsedDay) &&
                        TryCreateDate(anchor, parsedDay, out date))
                    {
                        return true;
                    }
                }

                if (value is IFormattable formattable)
                {
                    var formatted = formattable.ToString(null, CultureInfo.InvariantCulture);
                    if (int.TryParse(formatted, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsedDay) &&
                        TryCreateDate(anchor, parsedDay, out date))
                    {
                        return true;
                    }
                }
            }

            date = default;
            return false;
        }
        private static bool TryCreateDate(DateTime anchor, int day, out DateTime result)
        {
            if (day <= 0)
            {
                result = default;
                return false;
            }

            var y = anchor.Year;
            var m = anchor.Month;
            var clampedDay = Math.Clamp(day, 1, DateTime.DaysInMonth(y, m));
            result = new DateTime(y, m, clampedDay);
            return true;
        }
        private static DateTime? FindDisplayMonth(object[] values)
        {
            if (values.Length > 2 && TryExtractDate(values[2], out var displayDate))
            {
                return new DateTime(displayDate.Year, displayDate.Month, 1);
            }
            for (var i = 0; i < values.Length; i++)
            {
                if (TryExtractDate(values[i], out var candidate))
                {
                    return new DateTime(candidate.Year, candidate.Month, 1);
                }
            }

            return null;
        }

        private static bool TryExtractDate(object? value, out DateTime date)
        {
            switch (value)
            {
                case DateTime dt:
                    date = dt; return true;
                case DateTimeOffset off:
                    date = off.Date; return true;
                case DateOnly d:
                    date = d.ToDateTime(TimeOnly.MinValue); return true;
                case string s when DateTime.TryParse(s, CultureInfo.CurrentCulture,
                                                     DateTimeStyles.AssumeLocal, out var parsed):
                    date = parsed; return true;
                default:
                    date = default; return false;
            }
        }
    }
}