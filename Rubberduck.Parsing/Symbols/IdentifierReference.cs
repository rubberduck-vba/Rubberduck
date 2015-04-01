using System;
using Antlr4.Runtime;

namespace Rubberduck.Parsing.Symbols
{
    public class IdentifierReference
    {
        public IdentifierReference(QualifiedModuleName qualifiedName, string identifierName, 
            Selection selection, bool isAssignment, ParserRuleContext context, Declaration declaration)
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

        private readonly ParserRuleContext _context;
        public ParserRuleContext Context { get { return _context; } }

        private readonly Declaration _declaration;
        public Declaration Declaration { get { return _declaration; } }

        public bool HasTypeHint()
        {
            try
            {
                var hint = ((dynamic) Context).typeHint();
                return hint != null && !string.IsNullOrEmpty(hint.GetText());
            }
            catch (Exception)
            {
                return false;
                throw;
            }
        }
    }
}