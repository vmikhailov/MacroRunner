using System;

namespace MacroRunner.Runtime.Excel.Cells;

public class FunctionCell : BaseCell
{
    private bool _invalidated;
    private Func<IExecutionContext, object> _formula;
    private IExecutionContext _context;

    public FunctionCell(IExecutionContext context, Func<IExecutionContext, object> formula)
    {
        _context = context;
        _formula = formula;
        Invalidate();
    }

    public void Invalidate()
    {
        _invalidated = true;
    }
    
    public override object? GetValue()
    {
        EnsureCalculated();
        return base.GetValue();
    }

    protected override void SetValue(object value)
    {
        base.SetValue(value);
        _invalidated = false;
    }

    private void Calculate()
    {
        _context.LockCell(this);
        SetValue(_formula(_context));
        _context.UnlockCell(this);
    }

    private void EnsureCalculated()
    {
        if (_invalidated)
        {
            Calculate();
        }
    }
}