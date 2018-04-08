using System;

namespace Rubberduck.VBEditor
{
    public struct Selection : IEquatable<Selection>, IComparable<Selection>
    {
        public Selection(int line, int column) : this(line, column, line, column) { }

        public Selection(int startLine, int startColumn, int endLine, int endColumn)
        {
            StartLine = startLine;
            StartColumn = startColumn;
            EndLine = endLine;
            EndColumn = endColumn;
        }

        /// <summary>
        /// The first character of the first line.
        /// </summary>
        public static Selection Home => new Selection(1, 1, 1, 1);

        public static Selection Empty => new Selection();

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

        public int StartLine { get; }

        public int EndLine { get; }

        public int StartColumn { get; }

        public int EndColumn { get; }

        public int LineCount => EndLine - StartLine + 1;

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
            return (StartLine == EndLine && StartColumn == EndColumn)
                ? string.Format(VBEEditorText.SelectionLocationPosition, StartLine, StartColumn)
                : string.Format(VBEEditorText.SelectionLocationRange, StartLine, StartColumn, EndLine, EndColumn);
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
