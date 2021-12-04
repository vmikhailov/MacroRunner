namespace MacroRunner.Runtime.Excel;

public class Address
{
    public int RowNumber { get; set; }

    public bool FixedRow { get; set; }

    public string ColumnName { get; set; }

    public bool FixedColumn { get; set; }

    private int _columnNumber = -1;
    public int ColumnNumber
    {
        get
        {
            if (_columnNumber == -1)
            {
                var n = 0;
                foreach (var c in ColumnName)
                {
                    n = n * 26 + (c - 'A' + 1);
                }

                _columnNumber = n;
            }

            return _columnNumber;
        }
    }
}