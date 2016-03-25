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
        void ParseComponent(VBComponent vbComponent, TokenStreamRewriter rewriter = null);
        Task ParseComponentAsync(VBComponent component);
    }
}