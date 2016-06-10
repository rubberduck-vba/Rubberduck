using System;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Parsing
{
    public interface IRubberduckParser : IDisposable
    {
        RubberduckParserState State { get; }
    }
}
