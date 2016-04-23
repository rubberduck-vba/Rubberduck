using Rubberduck.Parsing.VBA;

namespace Rubberduck.Parsing
{
    public interface IRubberduckParser
    {
        RubberduckParserState State { get; }
    }
}