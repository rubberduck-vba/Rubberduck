using Antlr4.Runtime;

namespace Rubberduck.Parsing.Symbols
{
    public class IdentifierReference
    {
        public IdentifierReference(QualifiedModuleName qualifiedName, string identifierName, 
            Selection selection, bool isAssignment, RuleContext context, Declaration declaration)
        {
            _qualifiedName = qualifiedName;
            _identifierName = identifierName;
            _selection = selection;
            _isAssignment = isAssignment;
            _context = context;
            _declaration = declaration;
        }

        private readonly QualifiedModuleName _qualifiedName;
        public QualifiedModuleName QualifiedModuleName { get { return _qualifiedName; } }

        private readonly string _identifierName;
        public string IdentifierName { get { return _identifierName; } }

        private readonly Selection _selection;
        public Selection Selection { get { return _selection; } }

        private readonly bool _isAssignment;
        public bool IsAssignment { get { return _isAssignment; } }

        private readonly RuleContext _context;
        public RuleContext Context { get { return _context; } }

        private readonly Declaration _declaration;
        public Declaration Declaration { get { return _declaration; } }
    }
}