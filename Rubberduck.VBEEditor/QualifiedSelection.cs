using System;

namespace Rubberduck.VBEditor
{
    public struct QualifiedSelection : IComparable<QualifiedSelection>, IEquatable<QualifiedSelection>
    {
        public QualifiedSelection(QualifiedModuleName qualifiedName, Selection selection)
        {
            _qualifiedName = qualifiedName;
            _selection = selection;
        }

        private readonly QualifiedModuleName _qualifiedName;
        public QualifiedModuleName QualifiedName { get { return _qualifiedName; } }

        private readonly Selection _selection;
        public Selection Selection { get { return _selection; } }

        public int CompareTo(QualifiedSelection other)
        {
            if (other.QualifiedName != QualifiedName)
            {
                return string.Compare(QualifiedName.ToString(), other.QualifiedName.ToString(), StringComparison.Ordinal);
            }

            return Selection.CompareTo(other.Selection);
        }

        public bool Equals(QualifiedSelection other)
        {
            return other.Selection.Equals(_selection) && other.QualifiedName.Equals(_qualifiedName);
        }

        public override string ToString()
        {
            return string.Concat(QualifiedName, " ", Selection);
        }

        public override int GetHashCode()
        {
            return HashCode.Compute(_qualifiedName.GetHashCode(), _selection.GetHashCode());
        }

        public static bool operator ==(QualifiedSelection selection1, QualifiedSelection selection2)
        {
            return selection1.Equals(selection2);
        }

        public static bool operator !=(QualifiedSelection selection1, QualifiedSelection selection2)
        {
            return !(selection1 == selection2);
        }

        public override bool Equals(object obj)
        {
            if (obj is QualifiedSelection)
            {
                return Equals((QualifiedSelection) obj);
            }
            return false;
        }
    }
}
