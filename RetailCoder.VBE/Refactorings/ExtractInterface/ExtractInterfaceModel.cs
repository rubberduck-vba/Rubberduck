using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.ExtractInterface
{
    public class ExtractInterfaceModel
    {
        private readonly RubberduckParserState _parseResult;
        public RubberduckParserState ParseResult { get { return _parseResult; } }

        private readonly IEnumerable<Declaration> _declarations;
        public IEnumerable<Declaration> Declarations { get { return _declarations; } }

        private readonly QualifiedSelection _selection;
        public QualifiedSelection Selection { get { return _selection; } }

        private readonly Declaration _targetDeclaration;
        public Declaration TargetDeclaration { get { return _targetDeclaration; } }

        private readonly List<InterfaceMember> _members;
        public List<InterfaceMember> Members { get { return _members; } }

        private readonly static DeclarationType[] DeclarationTypes =
        {
            DeclarationType.Class,
            DeclarationType.Document,
            DeclarationType.UserForm
        };

        public ExtractInterfaceModel(RubberduckParserState parseResult, QualifiedSelection selection)
        {
            _parseResult = parseResult;
            _selection = selection;
            _declarations = parseResult.AllDeclarations.ToList();

            _targetDeclaration =
                _declarations.SingleOrDefault(
                    item =>
                        !item.IsBuiltIn && DeclarationTypes.Contains(item.DeclarationType)
                        && item.Project == selection.QualifiedName.Project
                        && item.QualifiedSelection.QualifiedName == selection.QualifiedName);

             _members = _declarations.Where(item => !item.IsBuiltIn &&
                                                item.Project == _targetDeclaration.Project &&
                                                item.ComponentName == _targetDeclaration.ComponentName &&
                                                item.Accessibility == Accessibility.Public)
                                     .OrderBy(o => o.Selection.StartLine)
                                     .ThenBy(t => t.Selection.StartColumn)
                                     .Select(d => new InterfaceMember(d, _declarations))
                                     .ToList();
        }
    }
}