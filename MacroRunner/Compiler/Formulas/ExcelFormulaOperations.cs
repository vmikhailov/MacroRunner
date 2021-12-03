namespace MacroRunner.Compiler.Formulas
{
    public class ExcelFormulaOperations
    {
        public static int Add(int a, int b) => a + b;
        
        public static double Add(double a, double b) => a + b;

        public static dynamic Add(dynamic a, dynamic b) => a + b;

        public static int Subtract(int a, int b) => a - b;

        public static double Subtract(double a, double b) => a - b;

        public static int Multiply(int a, int b) => a * b;

        public static double Multiply(double a, double b) => a * b;

        public static double Divide(double a, double b) => a / b;

        public static double Modulo(double a, double b) => a % b;
    }
}