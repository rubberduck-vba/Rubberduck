using System;
using Antlr4.Runtime;

namespace Rubberduck.Parsing.VBA.Parsing.ParsingExceptions
{
    public interface IRubberduckParseErrorListener : IParserErrorListener
    {
        bool HasPostponedException(out Exception exception);
    }
}
