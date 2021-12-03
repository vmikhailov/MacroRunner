namespace MacroRunner.Runtime.Excel
{
    public class Workbooks
    {
        public Range Cells { get; set; }

        public Workbook this[string name] => null;

        public Workbook Open(string name)
        {
            return null;
        }
    }
}