using System;
using System.Collections.Generic;
using MacroRunner.Runtime.Excel.Cells;

namespace MacroRunner.Runtime;

public class FormulaExecutionContext : IExecutionContext
{
    private IDictionary<string, object?> _values = new Dictionary<string, object?>();
    private IExecutionContext? _parent;

    public FormulaExecutionContext()
    {
    }

    private FormulaExecutionContext(IExecutionContext parent)
    {
        _parent = parent;
    }

    public bool Root => _parent == null;

    public IExecutionContext Scoped => new FormulaExecutionContext(this);

    public IExecutionContext GetScopedWithValue(string name, object? value)
    {
        var ctx = Scoped;
        ctx.SetNamedValue(name, value);
        return ctx;
    }

    public ICell GetNamedCell(string name)
    {
        throw new System.NotImplementedException();
    }

    public object? GetNamedValue(string name)
    {
        if (_values.TryGetValue(name, out var value))
        {
            return value;
        }

        if (_parent == null)
        {
            throw new($"Value for {name} is missing in the context.");
        }
        
        return _parent.GetNamedValue(name);
    }

    public void SetNamedValue(string name, object? value)
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

    public object? this[string name]
    {
        get => GetNamedValue(name);
        set => SetNamedValue(name, value);
    }
}