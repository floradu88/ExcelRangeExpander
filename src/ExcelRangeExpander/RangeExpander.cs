using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelRangeExpander.Helpers;
using ExcelRangeExpander.Interfaces;

namespace ExcelRangeExpander
{
    public class RangeExpander : IRangeExpander
    {
        private readonly int maximumRowCount = int.Parse(ConfigurationManager.AppSettings["ExcelRangeExpander.MaximumRowCount"] ?? "65535");

        public string Expand(string range)
        {
            var result = string.Empty;
            var ranges = new List<string>();

            if (string.IsNullOrEmpty(range))
                return result;

            extractRanges(range, ref ranges);

            if (ranges.Any())
                result = string.Join(",", ranges);

            return result;
        }

        public List<string> ExpandList(string range)
        {
            var ranges = new List<string>();

            if (string.IsNullOrEmpty(range))
                return ranges;

            extractRanges(range, ref ranges);

            return ranges;
        }

        private void extractRanges(string range, ref List<string> ranges)
        {
            var rangeValues = range.Split(new string[] { "," }, StringSplitOptions.None);

            foreach (var rangeValue in rangeValues)
            {
                ranges.AddRange(Parse(rangeValue));
            }
        }

        private List<string> Parse(string range)
        {
            var values = range.Split(new string[] { ":" }, StringSplitOptions.RemoveEmptyEntries);
            var result = new List<string>();

            var startValue = values.FirstOrDefault();
            string endValue = "";

            if (values.Length == 2)
            {
                endValue = values.LastOrDefault();
            }

            var startValueString = GetResultString(startValue, @"\d+");
            var endValueString = GetResultString(endValue, @"\d+");
            var startColumnString = Regex.Replace(startValue, @"\d+", string.Empty);
            var endColumnString = Regex.Replace(endValue, @"\d+", string.Empty);

            if (string.IsNullOrEmpty(endValue))
                endColumnString = startColumnString;

            int start = 0;
            int end = 0;
            int.TryParse(startValueString, out start);
            int.TryParse(endValueString, out end);

            if (start != 0 && end != 0 && start <= end)
            {
                ExpandRange(result, startColumnString, endColumnString, start, end);
            }
            else if (start == 0 && end == 0)
            {
                start = 1;
                end = maximumRowCount;
                ExpandRange(result, startColumnString, endColumnString, start, end);
            }
            else if (end == 0 && start != 0)
            {
                end = maximumRowCount;

                ExpandRange(result, startColumnString, endColumnString, start, end);
            }
            else
            {
                throw new InvalidOperationException("Invalid range");
            }

            return result;
        }

        private static void ExpandRange(List<string> result, string startColumnString, string endColumnString, int start, int end)
        {
            if (startColumnString == endColumnString)
            {
                for (int i = start; i <= end; i++)
                {
                    result.Add(startColumnString + i);
                }
            }
            else
            {
                var endRange = ExcelHelper.ColumnNameToNumber(endColumnString);
                var startRange = ExcelHelper.ColumnNameToNumber(startColumnString);

                for (int c = startRange; c <= endRange; c++)
                {
                    var columnName = ExcelHelper.NumberToColumnName(c);
                    for (int i = start; i <= end; i++)
                    {
                        result.Add(columnName + i);
                    }
                }
            }
        }

        private static string GetResultString(string startValue, string pattern)
        {
            return Regex.Match(startValue, pattern).Value;
        }
    }
}
