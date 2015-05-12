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

        public string GetLines(Selection selection)
        {
            return Editor.get_Lines(selection.StartLine, selection.LineCount);
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