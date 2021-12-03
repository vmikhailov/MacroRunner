using System;

namespace MacroRunner.Runtime.Excel
{
    public class Range
    {
        public Address From;
        public Address To;

        public Range()
        {
        }

        public Range(Address address)
        {
            From = address;
            To = address;
        }

        public Range(Address from, Address to)
        {
            From = @from;
            To = to;
        }

        public Range Cells { get; set; }

        public int Column { get; set; }

        public Range Columns { get; set; }

        public int Count { get; }

        public object FormulaR1C1 { get; set; }

        public Range this[int row, int column]
        {
            get
            {
                var r = new Range();
                r.Value = 1;
                return r;
            }
            set { }
        }

        public Range this[long row, long column] => null;

        /*
        public Range this[Range from, Range to] => null;

        public Range this[string rangeName] => null;
        */
        public Range this[params object[] objs] => null;

        public int Row { get; set; }

        public Range Rows { get; set; }

        public object Value { get; set; }

        public object Value2 { get; set; }

        public static implicit operator string(Range r)
        {
            return (string)Convert.ChangeType(r.Value, TypeCode.String);
        }

        public static implicit operator int(Range r)
        {
            return (int)Convert.ChangeType(r.Value, TypeCode.Int32);
        }

        public static implicit operator double(Range r)
        {
            return (double)Convert.ChangeType(r.Value, TypeCode.Double);
        }

        public static implicit operator decimal(Range r)
        {
            return (decimal)Convert.ChangeType(r.Value, TypeCode.Decimal);
        }

        public void AutoFilter()
        {
        }

        public void Clear()
        {
        }

        public void ClearComments()
        {
        }

        public void ClearContents()
        {
        }

        public void ClearFormats()
        {
        }

        public void ClearHyperlinks()
        {
        }

        public void Group()
        {
        }

        public void Select()
        {
        }

        public void Ungroup()
        {
        }
    }
}