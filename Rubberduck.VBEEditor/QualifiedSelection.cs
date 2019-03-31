using System;

namespace Rubberduck.VBEditor
{
    public readonly struct QualifiedSelection : IComparable<QualifiedSelection>, IEquatable<QualifiedSelection>
    {
        public QualifiedSelection(QualifiedModuleName qualifiedName, Selection selection)
        {
            QualifiedName = qualifiedName;
            Selection = selection;
        }

        public QualifiedModuleName QualifiedName { get; }

        public Selection Selection { get; }

        public bool Contains(QualifiedSelection other)
        {
            return QualifiedName.Equals(other.QualifiedName) && Selection.Contains(other.Selection);
        }

        public int CompareTo(QualifiedSelection other)
        {
            return other.QualifiedName == QualifiedName
                ? Selection.CompareTo(other.Selection)
                : string.Compare(QualifiedName.ToString(), other.QualifiedName.ToString(), StringComparison.Ordinal);
        }

        public bool Equals(QualifiedSelection other)
        {
            return other.Selection.Equals(Selection) && other.QualifiedName.Equals(QualifiedName);
        }

        public override string ToString()
        {
            return $"{QualifiedName} {Selection}";
        }

        public override int GetHashCode()
        {
            return HashCode.Compute(QualifiedName.GetHashCode(), Selection.GetHashCode());
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
            if (obj is QualifiedSelection qualifiedSelection)
            {
                return Equals(qualifiedSelection);
            }
            return false;
        }
    }
}
