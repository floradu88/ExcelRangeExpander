using ExcelRangeExpander.Helpers;
using Xunit;

namespace ExcelRangeExpander.Tests
{
    public class ExcelHelperTests
    {
        [Theory]
        [InlineData(1, "A")]
        [InlineData(2, "B")]
        [InlineData(26, "Z")]
        [InlineData(27, "AA")]
        [InlineData(54, "BB")]
        public void should_convert_number_to_excel_column_name(int number, string columnName)
        {
            Assert.Equal(ExcelHelper.NumberToColumnName(number), columnName);
        }

        [Theory]
        [InlineData("A", 1)]
        [InlineData("B", 2)]
        [InlineData("Z", 26)]
        [InlineData("AA", 27)]
        [InlineData("BB", 54)]
        public void should_convert_excel_column_name_to_number(string columnName, int number)
        {
            Assert.Equal(ExcelHelper.ColumnNameToNumber(columnName), number);
        }
    }
}
