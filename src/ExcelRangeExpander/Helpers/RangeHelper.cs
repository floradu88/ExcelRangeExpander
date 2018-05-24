using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelRangeExpander.Helpers
{
    public static class RangeHelper
    {
        private const string Pattern = @"\d+";

        public static bool ParseRange(string range, out string startColumn, out string endColumn, out int start,
            out int end)
        {
            startColumn = null;
            endColumn = null;
            start = 0;
            end = 0;

            if (string.IsNullOrEmpty(range))
                return false;

            var values = range.Split(new string[] { ":" }, StringSplitOptions.RemoveEmptyEntries);
            var result = new List<string>();

            var startValue = values.FirstOrDefault();
            string endValue = "";

            if (values.Length == 2)
            {
                endValue = values.LastOrDefault();
            }

            var startValueString = GetResultString(startValue, Pattern);
            startColumn = ExtractColumnName(startValue);
            var startColumnString = ExcelHelper.ColumnNameToNumber(startColumn);

            if (string.IsNullOrEmpty(endValue))
                endValue = startColumn;

            var endValueString = GetResultString(endValue, Pattern);
            endColumn = ExtractColumnName(endValue);

            var endColumnString = ExcelHelper.ColumnNameToNumber(endColumn);

            if (string.IsNullOrEmpty(endValue))
                endColumnString = startColumnString;
            if (!int.TryParse(startValueString, out start))
                start = 1;
            int.TryParse(endValueString, out end);

            return true;
        }

        public static bool IsInRange(string range, string column)
        {
            int start = 0;
            int end = 0;
            string startColumn, endColumn;

            if (!ParseRange(range, out startColumn, out endColumn, out start, out end))
                return false;

            var valueString = RangeHelper.GetResultString(column, Pattern);
            var columnString = ExcelHelper.ColumnNameToNumber(RangeHelper.ExtractColumnName(column));

            int current = 0;
            int.TryParse(valueString, out current);

            if (ExcelHelper.ColumnNameToNumber(startColumn) <= columnString && ExcelHelper.ColumnNameToNumber(endColumn) >= columnString)
            {
                if (current >= start && current <= end)
                {
                    return true;
                }
            }

            return false;
        }

        public static string GetResultString(string startValue, string pattern)
        {
            return Regex.Match(startValue, pattern).Value;
        }

        public static string ExtractColumnName(string value)
        {
            return Regex.Replace(value, Pattern, string.Empty);
        }
    }
}
