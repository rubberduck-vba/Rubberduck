namespace Rubberduck.SmartIndenter
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

        private readonly int _startLine;
        public int StartLine { get { return _startLine; } }

        private readonly int _startColumn;
        public int StartColumn { get { return _startColumn; } }

        private readonly int _endLine;
        public int EndLine { get { return _endLine; } }

        private readonly int _endColumn;
        public int EndColumn { get { return _endColumn; } }

        public int LineCount { get { return _endLine - _startLine + 1; } }
    }
}
