using System;

namespace MacroRunner.Runtime.Excel;

public class ExcelDate
{
    public static ExcelDate? operator -(ExcelDate d1, ExcelDate d2)
    {
        return null;
    }

    public static implicit operator DateTime(ExcelDate d)
    {
        return DateTime.Now;
    }
}