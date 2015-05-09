namespace Rubberduck.VBEditor
{
    public struct Selection
    {
        public Selection(int startLine, int startColumn, int endLine, int endColumn)
        {
            _startLine = startLine;
            _startColumn = startColumn;
            _endLine = endLine;
            _endColumn = endColumn;
        }

        public static Selection Home { get { return new Selection(1, 1, 1, 1); } }

        public bool ContainsFirstCharacter(Selection selection)
        {
            return Contains(new Selection(selection.StartLine, selection.StartColumn, selection.StartLine, selection.StartColumn));
        }

        public bool Contains(Selection selection)
        {
            // single line comparison
            if (selection.StartLine == StartLine && selection.EndLine == EndLine)
            {
                return selection.StartColumn >= StartColumn && selection.EndColumn <= EndColumn;
            }

            // multiline, obvious case:
            if (selection.StartLine > StartLine && selection.EndLine < EndLine)
            {
                return true;
            }

            // starts on same line:
            if (selection.StartLine == StartLine && selection.StartColumn > StartColumn)
            {
                return selection.EndLine < EndLine || 
                    (selection.EndLine == EndLine && selection.EndColumn <= EndColumn);
            }

            // ends on same line:
            if (selection.EndLine == EndLine && selection.EndColumn < EndColumn)
            {
                return selection.StartLine > StartLine ||
                       (selection.StartLine == StartLine && selection.StartColumn >= StartColumn);
            }

            return false;
        }

        private readonly int _startLine;
        public int StartLine { get { return _startLine; } }
        
        private readonly int _endLine;
        public int EndLine { get { return _endLine; } }

        private readonly int _startColumn;
        public int StartColumn { get { return _startColumn; } }
        
        private readonly int _endColumn;
        public int EndColumn { get { return _endColumn; } }

        public int LineCount { get { return _endLine - _startLine + 1; } }

        public override string ToString()
        {
            return string.Format("Start: L{0}C{1} End: L{2}C{3}", _startLine, _startColumn, _endLine, _endColumn);
        }
    }
}