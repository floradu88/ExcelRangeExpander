using System;
using System.Collections.Generic;
using System.Text;
using Xunit;

namespace ExcelRangeExpander.Tests
{
    public class RangeExpanderTests
    {
        [Theory]
        [InlineData("A1:A2", "A1,A2")]
        [InlineData("A1:A3", "A1,A2,A3")]
        [InlineData("A1:A1", "A1")]
        [InlineData("A1:B2", "A1,A2,B1,B2")]
        [InlineData("A1:C3", "A1,A2,A3,B1,B2,B3,C1,C2,C3")]
        
        public void should_expand_a_one_column_range_correctly(string range, string expectedRange)
        {
            IRangeExpander rangeExpander = new RangeExpander();

            var actualRange = rangeExpander.Expand(range);

            Assert.Equal(expectedRange, actualRange);
        }

        [Theory]
        [InlineData("A1:A")]
        [InlineData("A1:")]
        [InlineData("A3:C1")]
        public void should_throw_exception_if_range_is_invalid(string range)
        {
            IRangeExpander rangeExpander = new RangeExpander();

            Assert.Throws<InvalidOperationException>(() => rangeExpander.Expand(range));
        }

        [Fact]
        public void should_full_expand_a_one_column_range_correctly()
        {
            string range = "A:A";
            string expectedRange = "";
            var rangeValues = new List<string>();

            for (int i = 1; i <= 65535; i++)
            {
                rangeValues.Add("A" + i);
            }

            expectedRange = string.Join(",", rangeValues);

            IRangeExpander rangeExpander = new RangeExpander();

            var actualRange = rangeExpander.Expand(range);

            Assert.Equal(expectedRange, actualRange);
        }
    }
}
