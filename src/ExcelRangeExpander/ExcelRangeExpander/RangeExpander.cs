using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

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

            if (values.Length != 2)
            {
                throw new InvalidOperationException("Invalid range");
            }

            var startValue = values.First();
            var endValue = values.Last();

            var startValueString = GetResultString(startValue, @"\d+");
            var endValueString = GetResultString(endValue, @"\d+");
            var startColumnString = Regex.Replace(startValue, @"\d+", string.Empty);
            var endColumnString = Regex.Replace(endValue, @"\d+", string.Empty);

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
                for (char c = startColumnString.First(); c <= endColumnString.First(); c++)
                {
                    for (int i = start; i <= end; i++)
                    {
                        result.Add("" + c + i);
                    }
                }
            }
        }

        private static string GetResultString(string startValue, string pattern)
        {
            return Regex.Match(startValue, pattern).Value;
        }
    }

    public interface IRangeExpander
    {
        string Expand(string range);
    }
}
