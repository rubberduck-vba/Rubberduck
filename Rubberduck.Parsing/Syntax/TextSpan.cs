using System;

namespace Rubberduck.Parsing.Syntax
{
    public struct TextSpan : IEquatable<TextSpan>, IComparable<TextSpan>
    {
        public const int BaseLine = 1;
        public const int BaseColumn = 1;

        public TextSpan(int startLine, int startColumn, int endLine, int endColumn)
        {
            if (startColumn < BaseColumn || startLine < BaseLine) { throw new ArgumentException(); }

            _startLine = startLine;
            _startColumn = startColumn;
            _endLine = endLine;
            _endColumn = endColumn;
        }

        public static TextSpan Expand(TextSpan span, TextSpan child)
        {
            return new TextSpan(span.StartLine, span.StartColumn, child.EndLine, child.EndColumn);
        }

        private readonly int _startLine;
        public int StartLine { get { return _startLine; } }

        private readonly int _startColumn;
        public int StartColumn { get { return _startColumn; } }

        private readonly int _endLine;
        public int EndLine { get { return _endLine; } }

        private readonly int _endColumn;
        public int EndColumn { get { return _endColumn; } }

        public int Lines { get { return _endLine - _startLine + 1; } }
        public int Columns { get { return _endColumn - _startColumn + 1; } }

        public bool IsEmpty { get { return _startLine == _endLine && _startColumn == _endColumn; } }

        public bool Contains(TextSpan other)
        {
            return other.StartLine >= _startLine 
                && other.StartColumn >= _startColumn
                && other.EndLine <= _endLine
                && other.EndColumn <= _endColumn;
        }

        public bool Overlaps(TextSpan other)
        {
            var startLineOverlap = Math.Max(_startLine, other.StartLine);
            var endLineOverlap = Math.Min(_endLine, other.EndLine);

            var startColumnOverlap = Math.Max(_startColumn, other.StartColumn);
            var endColumnOverlap = Math.Min(_endColumn, other.EndColumn);

            return (startLineOverlap < endLineOverlap) 
                && (startColumnOverlap < endColumnOverlap);
        }

        public TextSpan? Overlap(TextSpan other)
        {
            var startLineOverlap = Math.Max(_startLine, other.StartLine);
            var endLineOverlap = Math.Min(_endLine, other.EndLine);

            var startColumnOverlap = Math.Max(_startColumn, other.StartColumn);
            var endColumnOverlap = Math.Min(_endColumn, other.EndColumn);

            var overlaps = (startLineOverlap < endLineOverlap)
                        && (startColumnOverlap < endColumnOverlap);
            return overlaps
                ? new TextSpan(startLineOverlap, startColumnOverlap, endLineOverlap, endColumnOverlap)
                : (TextSpan?)null;
        }

        public bool Intersects(TextSpan other)
        {
            return other.StartLine <= _endLine
                && other.StartColumn <= _endColumn
                && other.EndLine >= _startLine
                && other.EndColumn >= _startColumn;
        }

        public TextSpan? Intersection(TextSpan other)
        {
            var startLineIntersect = Math.Max(_startLine, other.StartLine);
            var endLineIntersect = Math.Min(_endLine, other.EndLine);

            var startColumnIntersect = Math.Max(_startColumn, other.StartColumn);
            var endColumnIntersect = Math.Min(_endColumn, other.EndColumn);

            var intersects = startLineIntersect <= endLineIntersect
                          && startColumnIntersect <= endColumnIntersect;
            return intersects
                ? new TextSpan(startLineIntersect, startColumnIntersect, endLineIntersect, endColumnIntersect) 
                : (TextSpan?)null;
        }

        public static bool operator ==(TextSpan left, TextSpan right)
        {
            return left.Equals(right);
        }

        public static bool operator !=(TextSpan left, TextSpan right)
        {
            return !(left == right);
        }

        public bool Equals(TextSpan other)
        {
            return _startLine == other.StartLine
                && _startColumn == other.StartColumn
                && _endLine == other.EndLine
                && _endColumn == other.EndColumn;
        }

        public int CompareTo(TextSpan other)
        {
            var lineDiff = _startLine - other.StartLine;
            if (lineDiff != 0)
            {
                return lineDiff;
            }

            var columnDiff = _endColumn - other.EndColumn;
            if (columnDiff != 0)
            {
                return columnDiff;
            }

            return Columns - other.Columns;
        }

        public override bool Equals(object obj)
        {
            return obj is TextSpan && Equals((TextSpan) obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hash = 17;
                hash = hash * 23 + _startLine.GetHashCode();
                hash = hash * 23 + _startColumn.GetHashCode();
                hash = hash * 23 + _endLine.GetHashCode();
                hash = hash * 23 + _endColumn.GetHashCode();
                return hash;
            }
        }

        public override string ToString()
        {
            return string.Format("{0}:{1}-{2}:{3}", _startLine, _startColumn, _endLine, _endColumn);
        }
    }
}