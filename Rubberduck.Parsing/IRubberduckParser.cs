using System.Threading.Tasks;
using Microsoft.Vbe.Interop;

namespace Rubberduck.Parsing
{
    public interface IRubberduckParser
    {
        /// <summary>
        /// Parses all code modules in specified project.
        /// </summary>
        /// <returns>Returns an <c>IParseTree</c> for each code module in the project; the qualified module name being the key.</returns>
        VBProjectParseResult Parse(VBProject vbProject);

        Task<VBProjectParseResult> ParseAsync(VBProject vbProject);

        void RemoveProject(VBProject vbProject);
    }
}