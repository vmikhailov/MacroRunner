using MacroRunner.Runtime.Excel.Cells;

namespace MacroRunner.Runtime.Excel;

public interface IExecutionContext
{
    ICell GetNamedCell(string name);

    void LockCell(ICell cell);

    void UnlockCell(ICell cell);
}