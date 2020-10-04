using System;
using MemoryStream = System.IO.MemoryStream;
using Stream = System.IO.Stream;
using StreamWriter = System.IO.StreamWriter;

namespace Rubberduck.VBEditor.Extensions
{
    public static class StringExtensions
    {
        public static bool Contains(this string source, string toCheck, StringComparison comp)
        {
            return source.IndexOf(toCheck, comp) >= 0;
        }

        public static Stream ToStream(this string source)
        {
            var stream = new MemoryStream();
            var writer = new StreamWriter(stream); // do not dispose the writer http://stackoverflow.com/a/1879470/1188513
            writer.Write(source);
            writer.Flush();
            stream.Position = 0;
            return stream;
        }
    }
}
