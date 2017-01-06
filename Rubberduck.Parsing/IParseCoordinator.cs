using System;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Parsing
{
    public interface IParseCoordinator : IDisposable
    {
        RubberduckParserState State { get; }
    }
}
