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

        public static bool GreaterThan(int a, int b) => a > b;

        public static bool GreaterThan(double a, double b) => a > b;

        public static bool LessThan(int a, int b) => a < b;

        public static bool LessThan(double a, double b) => a < b;

        public static bool Equal(int a, int b) => a == b;

        //excel compares to e-15
        public static bool Equal(double a, double b) => a == b;

        public static bool NotEqual(int a, int b) => a != b;

        public static bool NotEqual(double a, double b) => a != b;
    }
}