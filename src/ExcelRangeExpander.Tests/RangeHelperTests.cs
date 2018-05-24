using System;
using ExcelRangeExpander.Helpers;
using Xunit;

namespace ExcelRangeExpander.Tests
{
    public class RangeHelperTests
    {
        [Theory]
        [InlineData("A1:A100", "A1", true)]
        [InlineData("A1:A100", "A2", true)]
        [InlineData("A1:A100", "A100", true)]
        [InlineData("A1:B100", "A1", true)]
        [InlineData("A1:B100", "B1", true)]
        [InlineData("A1:A100", "B1", false)]
        [InlineData("A1:A100", "A101", false)]
        public void should_convert_number_to_excel_column_name(string range, string columnName, bool expectedOutcome)
        {
            Assert.Equal(RangeHelper.IsInRange(range, columnName), expectedOutcome);
        }
    }
}
