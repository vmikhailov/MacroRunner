using System.Collections.Generic;
using MacroRunner.Runtime.Excel.Cells;

namespace MacroRunner.Runtime;

public interface IExecutionContext
{
    ICell GetNamedCell(string name);

    object GetNamedValue(string name);

    void SetNamedValue(string name, object value);

    void LockCell(ICell cell);

    void UnlockCell(ICell cell);
}

public class FormulaExecutionContext : IExecutionContext
{
    private IDictionary<string, object> _values = new Dictionary<string, object>();
    public ICell GetNamedCell(string name)
    {
        throw new System.NotImplementedException();
    }

    public object GetNamedValue(string name)
    {
        return _values[name];
    }

    public void SetNamedValue(string name, object value)
    {
        _values[name] = value;
    }

    public void LockCell(ICell cell)
    {
        throw new System.NotImplementedException();
    }

    public void UnlockCell(ICell cell)
    {
        throw new System.NotImplementedException();
    }
}