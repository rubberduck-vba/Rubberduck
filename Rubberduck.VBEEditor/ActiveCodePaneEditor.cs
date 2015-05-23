using Rubberduck.VBEditor.Extensions;
using Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor
{
    public class ActiveCodePaneEditor : IActiveCodePaneEditor
    {
        private readonly VBE _vbe;

        public ActiveCodePaneEditor(VBE vbe)
        {
            _vbe = vbe;
        }

        private CodeModule Editor { get { return _vbe.ActiveCodePane.CodeModule; } }

        public QualifiedSelection GetSelection()
        {
            return Editor.CodePane.GetSelection();
        }

        public string GetLines(Selection selection)
        {
            return Editor.get_Lines(selection.StartLine, selection.LineCount);
        }

        public string GetSelectedProcedureScope(Selection selection)
        {
            var moduleName = Editor.Name;
            var projectName = Editor.Parent.Collection.Parent.Name;
            var parentScope = projectName + '.' + moduleName;

            vbext_ProcKind kind;
            var procStart = Editor.get_ProcOfLine(selection.StartLine, out kind);
            var procEnd = Editor.get_ProcOfLine(selection.EndLine, out kind);

            return procStart == procEnd
                ? parentScope + '.' + procStart
                : null;
        }

        public void DeleteLines(Selection selection)
        {
            Editor.DeleteLines(selection.StartLine, selection.LineCount);
        }

        public void ReplaceLine(int line, string content)
        {
            Editor.ReplaceLine(line, content);
        }

        public void InsertLines(int line, string content)
        {
            Editor.InsertLines(line, content);
        }
    }
}