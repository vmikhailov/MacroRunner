using System.Collections.Generic;
using System.Linq;
using MacroRunner.Runtime.Excel;

namespace MacroRunner.Runtime
{
    public class RangeFactory
    {
        public static Range Create(IEnumerable<Address> addresses)
        {
            var adr = addresses.ToArray();
            return adr.Length switch
            {
                1 => new Range(adr[0]),
                2 => new Range(adr[0], adr[1]),
                _ => null
            };
        }
    }
}