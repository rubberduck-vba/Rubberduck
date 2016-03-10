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
        Task ParseComponentAsync(VBComponent component, TokenStreamRewriter rewriter = null);
        //Task ParseAsync(VBComponent component, CancellationToken token, TokenStreamRewriter rewriter = null);
        //void Resolve(CancellationToken token);
    }
}