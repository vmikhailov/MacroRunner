using MacroRunner.Runtime.Excel.Cells;

namespace MacroRunner.Runtime;

public interface IExecutionContext
{
    bool Root { get; }

    IExecutionContext Scoped { get; }
    
    IExecutionContext GetScopedWithValue(string name, object value);
    
    ICell GetNamedCell(string name);

    object? GetNamedValue(string name);

    void SetNamedValue(string name, object? value);

    void LockCell(ICell cell);

    void UnlockCell(ICell cell);

    object? this[string name]
    {
        get;
        set;
    }
}