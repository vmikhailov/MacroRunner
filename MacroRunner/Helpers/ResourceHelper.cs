using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace MacroRunner.Helpers
{
    internal class ResourceHelper
    {
        public static string GetResourceAsString(string name)
        {
            using (var stream = GetResourceAsStream(name))
            using (var reader = new StreamReader(stream))
            {
                return reader.ReadToEnd();
            }
        }

        public static Stream GetResourceAsStream(string name)
        {
            var assembly = Assembly.GetExecutingAssembly();
            return assembly.GetManifestResourceStream(name);
        }

        public static IEnumerable<TextResource> GetResourcesByMask(string mask, Func<string, string> nameExtract)
        {
            var regex = new Regex(mask);
            var assembly = Assembly.GetExecutingAssembly();
            var resources = assembly.GetManifestResourceNames()
                                    .OrderBy(x => x)
                                    .Where(x => regex.IsMatch(x))
                                    .Select(x => new TextResource
                                    {
                                        Name = nameExtract(x),
                                        Text = GetResourceAsString(x)
                                    })
                                    .ToList();

            return resources;
        }

    }
}