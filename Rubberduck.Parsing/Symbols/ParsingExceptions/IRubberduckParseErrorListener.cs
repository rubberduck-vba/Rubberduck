using System;
using Antlr4.Runtime;

namespace Rubberduck.Parsing.Symbols.ParsingExceptions
{
    public interface IRubberduckParseErrorListener : IParserErrorListener
    {
        bool HasPostponedException(out Exception exception);
    }
}
