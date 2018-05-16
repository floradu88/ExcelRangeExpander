using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelRangeExpander.Helpers;
using ExcelRangeExpander.Interfaces;

namespace ExcelRangeExpander
{
    public class RangeExpander : IRangeExpander
    {
        public string Expand(string range)
        {
            var result = string.Empty;
            var ranges = new List<string>();

            if (string.IsNullOrEmpty(range))
                return result;

            var rangeValues = range.Split(new string[] { "," }, StringSplitOptions.None);

            foreach (var rangeValue in rangeValues)
            {
                ranges.AddRange(Parse(rangeValue));
            }

            if (ranges.Any())
                result = string.Join(",", ranges);

            return result;
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
                end = 65535;
                ExpandRange(result, startColumnString, endColumnString, start, end);
            }
            else if (end == 0 && start != 0)
            {
                end = 65535;

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
                for (int c = ExcelHelper.ColumnNameToNumber(startColumnString); c <= ExcelHelper.ColumnNameToNumber(endColumnString); c++)
                {
                    for (int i = start; i <= end; i++)
                    {
                        result.Add(ExcelHelper.NumberToColumnName(c) + i);
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
