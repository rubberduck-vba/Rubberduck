using System.Collections.Generic;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.VB.Abstract;

namespace Rubberduck.SmartIndenter
{
    public interface IIndenter
    {
        void IndentCurrentProcedure();
        void IndentCurrentModule();
        void IndentCurrentProject();
        void Indent(IVBComponent component);
        void Indent(IVBComponent component, Selection selection);
        IEnumerable<string> Indent(IEnumerable<string> lines);
        IEnumerable<string> Indent(IEnumerable<string> codeLines, bool forceTrailingNewLines);
    }
}
