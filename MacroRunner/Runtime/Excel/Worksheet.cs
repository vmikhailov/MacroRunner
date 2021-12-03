namespace MacroRunner.Runtime.Excel
{
    public class Worksheet : Range
    {
        public bool AutoFilterMode { get; set; }

        public Outline Outline { get; set; }

        public Range Range { get; set; }

        public Range UsedRange { get; set; }

        public void Activate()
        {
        }
    }
}