using System.Runtime.InteropServices;
using Rubberduck.Inspections;

namespace Rubberduck.Extensions
{
    public struct QualifiedSelection
    {
        public QualifiedSelection(QualifiedModuleName qualifiedName, Selection selection)
        {
            _qualifiedName = qualifiedName;
            _selection = selection;
        }

        private readonly QualifiedModuleName _qualifiedName;
        public QualifiedModuleName QualifiedName { get {return _qualifiedName; } }

        private readonly Selection _selection;
        public Selection Selection { get { return Selection; } }
    }

    [ComVisible(false)]
    public struct Selection
    {
        public Selection(int startLine, int startColumn, int endLine, int endColumn)
        {
            _startLine = startLine;
            _startColumn = startColumn;
            _endLine = endLine;
            _endColumn = endColumn;
        }

        public static Selection Empty { get { return new Selection(1, 1, 1, 1); } }

        private readonly int _startLine;
        public int StartLine { get { return _startLine; } }
        
        private readonly int _endLine;
        public int EndLine { get { return _endLine; } }

        private readonly int _startColumn;
        public int StartColumn { get { return _startColumn; } }
        
        private readonly int _endColumn;
        public int EndColumn { get { return _endColumn; } }

        public int LineCount { get { return _endLine - _startLine + 1; } }
    }
}