using MacroRunner.Runtime.Excel;

namespace MacroRunner.Runtime;

public class ExcelGlobals
{
    public ExcelApplication Application { get; }

    public Range Range { get; }

    public Workbooks Workbooks { get; }

    public Workbook ActiveWorkbook { get; }

    public Worksheet ActiveSheet { get; }

    public Range Rows { get; }

    public Range Columns { get; }

    public Range UsedRange { get; }

    public Range Cells { get; }

    public Range Selection { get; }

    public int TestInt { get; set; }

    public int TestValue { get; set; }
}