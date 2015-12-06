using System;
using Microsoft.Vbe.Interop;

namespace Rubberduck.SmartIndenter
{
    public interface IIndenter
    {
        event EventHandler ReportProgress;
        void IndentCurrentProcedure();
        void IndentCurrentModule();
        void Indent(VBProject project);
        void Indent(VBComponent module);
        void Indent(VBComponent module, string procedureName);
        void Indent(string[] lines, string moduleName, bool reportProgress = true, int linesAlreadyRebuilt = 0);
    }
}
