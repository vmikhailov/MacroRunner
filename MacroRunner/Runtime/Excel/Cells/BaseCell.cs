namespace MacroRunner.Runtime.Excel.Cells;

public class BaseCell : ICell
{
    private object? _value;

    protected virtual void SetValue(object value)
    {
        _value = value;
    }
    
    public virtual object? GetValue()
    {
        return _value;
    }
}