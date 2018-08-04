using Antlr4.Runtime;
using Antlr4.Runtime.Misc;

namespace Rubberduck.Parsing
{
    public static class TokenStreamExtensions
    {
        public static string GetText(this ITokenStream tokenStream, int startIndex, int stopIndex)
        {
            var interval = new Interval(startIndex, stopIndex);
            return tokenStream.GetText(interval);
        }
    }
}
