using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.ExtractInterface
{
    public class ExtractInterfaceModel
    {
        public RubberduckParserState State { get; }
        public Declaration TargetDeclaration { get; }

        public string InterfaceName { get; set; }

        public ObservableCollection<InterfaceMember> Members { get; set; } = new ObservableCollection<InterfaceMember>();

        public IEnumerable<InterfaceMember> SelectedMembers => Members.Where(m => m.IsSelected);

        private static readonly DeclarationType[] ModuleTypes =
        {
            DeclarationType.ClassModule,
            DeclarationType.Document,
            DeclarationType.UserForm
        };

        public static readonly DeclarationType[] MemberTypes =
        {
            DeclarationType.Procedure,
            DeclarationType.Function,
            DeclarationType.PropertyGet,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet,
        };

        public ExtractInterfaceModel(RubberduckParserState state, QualifiedSelection selection)
        {
            State = state;
            var declarations = state.AllUserDeclarations.ToList();
            var candidates = declarations.Where(item => ModuleTypes.Contains(item.DeclarationType)).ToList();

            TargetDeclaration = candidates.SingleOrDefault(item => 
                        item.QualifiedSelection.QualifiedName.Equals(selection.QualifiedName));

            if (TargetDeclaration == null)
            {
                return;
            }

            InterfaceName = $"I{TargetDeclaration.IdentifierName}";

            Members = new ObservableCollection<InterfaceMember>(declarations.Where(item =>
                    item.ProjectId == TargetDeclaration.ProjectId
                    && item.ComponentName == TargetDeclaration.ComponentName
                    && (item.Accessibility == Accessibility.Public || item.Accessibility == Accessibility.Implicit)
                    && MemberTypes.Contains(item.DeclarationType))
                .OrderBy(o => o.Selection.StartLine)
                .ThenBy(t => t.Selection.StartColumn)
                .Select(d => new InterfaceMember(d))
                .ToList());
        }
    }
}
