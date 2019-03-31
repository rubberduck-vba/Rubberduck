using System.Collections.Generic;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.SmartIndenter
{
    public interface IIndenter
    {
        void IndentCurrentProcedure();
        void IndentCurrentModule();
        void IndentCurrentProject();
        void Indent(IVBComponent component);
        IEnumerable<string> Indent(string code);
        IEnumerable<string> Indent(IEnumerable<string> lines);
        IEnumerable<string> Indent(IEnumerable<string> codeLines, bool forceTrailingNewLines);
    }
}
