using System;
using System.Collections.Generic;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.VBA;

namespace Rubberduck.SmartIndenter
{
    public interface IIndenter
    {
        event EventHandler<IndenterProgressEventArgs> ReportProgress;
        void IndentCurrentProcedure();
        void IndentCurrentModule();
        void Indent(VBComponent component, bool reportProgress = true, int linesAlreadyRebuilt = 0);
        void Indent(VBComponent component, string procedureName, Selection selection, bool reportProgress = true, int linesAlreadyRebuilt = 0);
        IEnumerable<string> Indent(IEnumerable<string> lines, string moduleName, bool reportProgress = true, int linesAlreadyRebuilt = 0);
    }
}
