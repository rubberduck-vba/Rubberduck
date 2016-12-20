using System.Collections.Generic;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.SmartIndenter
{
    public interface IIndenter
    {
        void IndentCurrentProcedure();
        void IndentCurrentModule();
        void Indent(IVBComponent component);
        void Indent(IVBComponent component, string procedureName, Selection selection);
        IEnumerable<string> Indent(IEnumerable<string> lines, string moduleName);
    }
}
