namespace MacroRunner.Compiler.Formulas
{
    public class ExcelFormulaOperations
    {
        public static int Add(int a, int b) => a + b;
        
        public static decimal Add(decimal a, decimal b) => a + b;

        //public static dynamic Add(dynamic a, dynamic b) => a + b;

        public static int Subtract(int a, int b) => a - b;

        public static decimal Subtract(decimal a, decimal b) => a - b;

        public static int Multiply(int a, int b) => a * b;

        public static decimal Multiply(decimal a, decimal b) => a * b;

        public static decimal Divide(decimal a, decimal b) => a / b;

        public static decimal Modulo(decimal a, decimal b) => a % b;
    }
}