using System.Collections;

namespace MacroRunner.Runtime.Common
{
    public class Dictionary : Hashtable
    {
        public bool Exists(object value)
        {
            return Contains(value);
        }
    }
}