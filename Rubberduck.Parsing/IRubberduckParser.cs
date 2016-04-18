using System;
using System.Threading;
using System.Threading.Tasks;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Parsing
{
    public interface IRubberduckParser
    {
        RubberduckParserState State { get; }
        void LoadComReference(Reference item);
        void UnloadComReference(Reference reference);
        void ParseComponent(VBComponent component, TokenStreamRewriter rewriter = null);
        Task ParseAsync(VBComponent component, CancellationToken token,  TokenStreamRewriter rewriter = null);
        void Cancel(VBComponent component = null);
        void Resolve(CancellationToken token);
    }
}