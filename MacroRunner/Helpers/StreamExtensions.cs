using System.IO;

namespace MacroRunner.Helpers;

public static class StreamExtensions
{
    public static byte[] ReadAllBytes(this Stream stream)
    {
        var length = stream.Length;
        var buffer = new byte[length];
        stream.Read(buffer, 0, (int)length);
        return buffer;
    }
}