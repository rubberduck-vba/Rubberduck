using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.ExtractInterface
{
    public class ExtractInterfaceModel
    {
        private readonly Declaration _targetDeclaration;
        public Declaration TargetDeclaration { get { return _targetDeclaration; } }

        public string InterfaceName { get; set; }

        private IEnumerable<InterfaceMember> _members = new List<InterfaceMember>();
        public IEnumerable<InterfaceMember> Members { get { return _members; } set { _members = value; } }

        private static readonly DeclarationType[] ModuleTypes =
        {
            DeclarationType.Class,
            DeclarationType.Document,
            DeclarationType.UserForm
        };

        private static readonly DeclarationType[] MemberTypes =
        {
            DeclarationType.Procedure,
            DeclarationType.Function,
            DeclarationType.PropertyGet,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet,
        };

        public ExtractInterfaceModel(RubberduckParserState state, QualifiedSelection selection)
        {
            var declarations = state.AllDeclarations.ToList();
            var candidates = declarations.Where(item => !item.IsBuiltIn && ModuleTypes.Contains(item.DeclarationType)).ToList();

            _targetDeclaration = candidates.SingleOrDefault(item => 
                        item.QualifiedSelection.QualifiedName.Equals(selection.QualifiedName));

            if (_targetDeclaration == null)
            {
                //throw new InvalidOperationException();
                return;
            }

            InterfaceName = "I" + TargetDeclaration.IdentifierName;

            _members = declarations.Where(item => !item.IsBuiltIn
                                                  && item.ProjectId == _targetDeclaration.ProjectId
                                                  && item.ComponentName == _targetDeclaration.ComponentName
                                                  && (item.Accessibility == Accessibility.Public || item.Accessibility == Accessibility.Implicit)
                                                  && MemberTypes.Contains(item.DeclarationType))
                                   .OrderBy(o => o.Selection.StartLine)
                                   .ThenBy(t => t.Selection.StartColumn)
                                   .Select(d => new InterfaceMember(d, declarations))
                                   .ToList();
        }
    }
}