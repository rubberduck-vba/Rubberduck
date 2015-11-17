using System.Threading;
using System.Threading.Tasks;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Parsing
{
    public interface IRubberduckParser
    {
        RubberduckParserState State { get; }
        Task ParseAsync(VBComponent component, CancellationToken token);
        void Resolve(CancellationToken token);
    }
}