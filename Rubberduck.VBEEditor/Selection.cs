using System;

namespace Rubberduck.VBEditor
{
    public struct Selection : IEquatable<Selection>
    {
        public Selection(int startLine, int startColumn, int endLine, int endColumn)
        {
            _startLine = startLine;
            _startColumn = startColumn;
            _endLine = endLine;
            _endColumn = endColumn;
        }

        /// <summary>
        /// The first character of the first line.
        /// </summary>
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
            if (selection.StartLine == StartLine && selection.StartColumn >= StartColumn)
            {
                return selection.EndLine < EndLine || 
                    (selection.EndLine == EndLine && selection.EndColumn <= EndColumn);
            }

            // ends on same line:
            if (selection.EndLine == EndLine && selection.EndColumn <= EndColumn + 1) // +1 for \r\n
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

        public bool Equals(Selection other)
        {
            return other.StartLine == StartLine
                   && other.EndLine == EndLine
                   && other.StartColumn == StartColumn
                   && other.EndColumn == EndColumn;
        }

        public override string ToString()
        {
            return string.Format(Rubberduck.VBEditor.VBEEditorText.SelectionLocationInfo, _startLine, _startColumn, _endLine, _endColumn);
        }

        public override bool Equals(object obj)
        {
            return Equals((Selection) obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hash = 17;
                hash = hash*23 + StartLine;
                hash = hash*23 + EndLine;
                hash = hash*23 + StartColumn;
                hash = hash*23 + EndColumn;
                return hash;
            }
        }
    }
}