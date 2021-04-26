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

        public Selection ToZeroBased() => 
            new Selection(StartLine - 1, StartColumn - 1, EndLine - 1, EndColumn - 1);
        public Selection ToOneBased() =>
            new Selection(StartLine + 1, StartColumn + 1, EndLine + 1, EndColumn + 1);

        public Selection ShiftRight(int positions = 1) =>
            new Selection(StartLine, StartColumn + positions, EndLine, EndColumn + positions);

        public Selection ShiftLeft(int positions = 1) =>
            new Selection(StartLine, Math.Max(1, StartColumn - positions), EndLine, Math.Max(1, StartColumn - positions));

        public Selection Collapse() =>
            new Selection(EndLine, EndColumn, EndLine, EndColumn);

        public bool IsEmpty()
        {
            return Equals(Empty);
        }

        public bool ContainsFirstCharacter(Selection selection)
        {
            return Contains(new Selection(selection.StartLine, selection.StartColumn, selection.StartLine, selection.StartColumn));
        }

        public Selection ExtendLeft(int positions = 1)
        {
            return new Selection(StartLine, Math.Max(StartColumn - positions, 1), EndLine, EndColumn);
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

        public bool IsSingleLine => StartLine == EndLine;
        public bool IsSingleCharacter => IsSingleLine && StartColumn == EndColumn;

        public Selection PreviousLine => StartLine == 1 ? Home : new Selection(StartLine - 1, 1);
        public Selection NextLine => new Selection(StartLine + 1, 1);

        /// <summary>
        /// Adds each corresponding element of the specified <c>Selection</c> value. Useful for offsetting with a zero-based <c>Selection</c>.
        /// </summary>
        public Selection Offset(Selection offset)
        {
            return new Selection(StartLine + offset.StartLine, StartColumn + offset.StartColumn, EndLine + offset.EndLine, EndColumn + offset.EndColumn);
        }

        public int StartLine { get; }

        public int EndLine { get; }

        public int StartColumn { get; }

        public int EndColumn { get; }

        public int LineCount => EndLine - StartLine + 1;

        public bool Equals(Selection other)
        {
            return IsSamePosition(other.StartLine, other.StartColumn, StartLine, StartColumn)
                   && IsSamePosition(other.EndLine, other.EndColumn, EndLine, EndColumn);
        }

        private static bool IsSamePosition(int line1, int column1, int line2, int column2)
        {
            return line1 == line2 && column1 == column2;
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

        /// <summary>
        /// Orders first by start position and then end position.
        /// </summary>
        public static bool operator >(Selection selection1, Selection selection2)
        {
            return IsGreaterPosition(selection1.StartLine, selection1.StartColumn, selection2.StartLine, selection2.StartColumn)
                || IsSamePosition(selection1.StartLine, selection1.StartColumn, selection2.StartLine, selection2.StartColumn)
                    && IsGreaterPosition(selection1.EndLine, selection1.EndColumn, selection2.EndLine, selection2.EndColumn);
        }

        private static bool IsGreaterPosition(int line1, int column1, int line2, int column2)
        {
            return line1 > line2 
                || line1 == line2 
                    && column1 > column2;
        }

        /// <summary>
        /// Orders first by start position and then end position.
        /// </summary>
        public static bool operator <(Selection selection1, Selection selection2)
        {
            return IsLesserPosition(selection1.StartLine, selection1.StartColumn, selection2.StartLine, selection2.StartColumn)
                || IsSamePosition(selection1.StartLine, selection1.StartColumn, selection2.StartLine, selection2.StartColumn)
                    && IsLesserPosition(selection1.EndLine, selection1.EndColumn, selection2.EndLine, selection2.EndColumn);
        }

        private static bool IsLesserPosition(int line1, int column1, int line2, int column2)
        {
            return line1 < line2
                || line1 == line2
                    && column1 < column2;
        }

        /// <summary>
        /// Orders first by start position and then end position.
        /// </summary>
        public static bool operator >=(Selection selection1, Selection selection2)
        {
            return !(selection1 < selection2);
        }

        /// <summary>
        /// Orders first by start position and then end position.
        /// </summary>
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
