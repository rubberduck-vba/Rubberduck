namespace Rubberduck.Parsing.Symbols
{
    public struct IdentifierReference
    {
        public IdentifierReference(string projectName, string componentName, string identifierName, 
            Selection selection, bool isAssignment)
        {
            _projectName = projectName;
            _componentName = componentName;
            _identifierName = identifierName;
            _selection = selection;
            _isAssignment = isAssignment;
        }

        private readonly string _projectName;
        public string ProjectName { get { return _projectName; } }

        private readonly string _componentName;
        public string ComponentName { get { return _componentName; } }

        private readonly string _identifierName;
        public string IdentifierName { get { return _identifierName; } }

        private readonly Selection _selection;
        public Selection Selection { get { return _selection; } }

        private readonly bool _isAssignment;
        public bool IsAssignment { get { return _isAssignment; } }
    }
}