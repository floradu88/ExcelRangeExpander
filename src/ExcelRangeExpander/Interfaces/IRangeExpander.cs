using System.Collections.Generic;

namespace ExcelRangeExpander.Interfaces
{
    public interface IRangeExpander
    {
        string Expand(string range);

        List<string> ExpandList(string range);
    }
}