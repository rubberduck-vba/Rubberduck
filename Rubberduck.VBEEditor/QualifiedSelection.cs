namespace Rubberduck.VBEditor
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
        public Selection Selection { get { return _selection; } }

        public override string ToString()
        {
            return string.Concat(QualifiedName, " ", Selection);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hash = 17;
                hash = hash*23 + _qualifiedName.GetHashCode();
                hash = hash * 23 + _selection.GetHashCode();
                return hash;
            } 
        }

        public override bool Equals(object obj)
        {
            var other = (QualifiedSelection) obj;
            return other.QualifiedName.Equals(_qualifiedName)
                   && other.Selection.Equals(_selection);
        }
    }
}