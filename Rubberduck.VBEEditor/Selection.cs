using System;

namespace Rubberduck.VBEditor
{
    public struct Selection : IEquatable<Selection>, IComparable<Selection>
    {
        public Selection(int line, int column) : this(line, column, line, column) { }

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
        public static Selection Home
        {
            get { return new Selection(1, 1, 1, 1); }
        }

        public static Selection Empty
        {
            get { return new Selection(); }
        }

        public bool IsEmpty()
        {
            return Equals(Empty);
        }

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

        public int CompareTo(Selection other)
        {
            if (this > other)
            {
                return 1;
            }
            if (this < other)
            {
                return -1;
            }

            return 0;
        }

        public override string ToString()
        {
            return (_startLine == _endLine && _startColumn == _endColumn)
                ? string.Format(VBEEditorText.SelectionLocationPosition, _startLine, _startColumn)
                : string.Format(VBEEditorText.SelectionLocationRange, _startLine, _startColumn, _endLine, _endColumn);
        }

        public static bool operator ==(Selection selection1, Selection selection2)
        {
            return selection1.Equals(selection2);
        }

        public static bool operator !=(Selection selection1, Selection selection2)
        {
            return !(selection1 == selection2);
        }

        public static bool operator >(Selection selection1, Selection selection2)
        {
            return selection1.StartLine > selection2.StartLine ||
                   selection1.StartLine == selection2.StartLine &&
                   selection1.StartColumn > selection2.StartColumn;
        }

        public static bool operator <(Selection selection1, Selection selection2)
        {
            return selection1.StartLine < selection2.StartLine ||
                   selection1.StartLine == selection2.StartLine &&
                   selection1.StartColumn < selection2.StartColumn;
        }

        public static bool operator >=(Selection selection1, Selection selection2)
        {
            return !(selection1 < selection2);
        }

        public static bool operator <=(Selection selection1, Selection selection2)
        {
            return !(selection1 > selection2);
        }

        public override bool Equals(object obj)
        {
            return obj != null && Equals((Selection)obj);
        }

        public override int GetHashCode()
        {
            return HashCode.Compute(StartLine, EndLine, StartColumn, EndColumn);
        }
    }
}
