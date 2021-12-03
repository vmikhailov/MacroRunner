using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace MacroRunner.Helpers
{
    public static class TypeExtensions
    {
        public static void SetStaticField(this Type type, string name, object value)
        {
            type.GetField(name)!.SetValue(null, value);
        }

        public static object? GetStaticField(this Type type, string name)
        {
            return type.GetField(name)!.GetValue(null);
        }

        public static IEnumerable<MethodInfo> GetPublicStaticMethods(this Type type, string name, int args = 0) =>
            type.GetTypeInfo()
                .GetMethods(BindingFlags.Static | BindingFlags.Public)
                .Where(x => StringComparer.InvariantCultureIgnoreCase.Compare(x.Name, name) == 0)
                .Where(x => args == 0 || x.GetParameters().Length == args);
    }
}