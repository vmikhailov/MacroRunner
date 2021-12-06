using System.Collections;

namespace MacroRunner.Runtime.Common;

public class Dictionary1 : Hashtable
{
    public bool Exists(object value)
    {
        return Contains(value);
    }
}