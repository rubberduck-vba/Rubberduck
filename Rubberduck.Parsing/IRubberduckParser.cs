using System;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Parsing
{
    public interface IRubberduckParser : IDisposable
    {
        RubberduckParserState State { get; }
        void Cancel(VBComponent component = null);
    }
}
