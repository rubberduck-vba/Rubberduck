namespace Rubberduck.Parsing
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
    }
}