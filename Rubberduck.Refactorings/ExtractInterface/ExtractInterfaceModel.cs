using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Refactorings.ExtractInterface
{
    public enum ClassInstancing
    {
        Private,
        Public
    }

    public class ExtractInterfaceModel : IRefactoringModel
    {
        public IDeclarationFinderProvider DeclarationFinderProvider { get; }

        public ClassModuleDeclaration TargetDeclaration { get; }
        public string InterfaceName { get; set; }
        public ObservableCollection<InterfaceMember> Members { get; set; } = new ObservableCollection<InterfaceMember>();
        public IEnumerable<InterfaceMember> SelectedMembers => Members.Where(m => m.IsSelected);
        public ClassInstancing InterfaceInstancing { get; set; }
        public ClassInstancing ImplementingClassInstancing => TargetDeclaration.IsExposed 
            ? ClassInstancing.Public 
            : ClassInstancing.Private;

        public static readonly DeclarationType[] MemberTypes =
        {
            DeclarationType.Procedure,
            DeclarationType.Function,
            DeclarationType.PropertyGet,
            DeclarationType.PropertyLet,
            DeclarationType.PropertySet,
        };

        public ExtractInterfaceModel(IDeclarationFinderProvider declarationFinderProvider, ClassModuleDeclaration target, ICodeBuilder codeBuilder)
        {
            TargetDeclaration = target;
            DeclarationFinderProvider = declarationFinderProvider;

            if (TargetDeclaration == null)
            {
                return;
            }

            InterfaceName = $"I{TargetDeclaration.IdentifierName}";
            InterfaceInstancing = ImplementingClassInstancing;

            LoadMembers(codeBuilder);
        }

        private void LoadMembers(ICodeBuilder codeBuilder)
        {
            Members = new ObservableCollection<InterfaceMember>(DeclarationFinderProvider.DeclarationFinder
                .Members(TargetDeclaration.QualifiedModuleName)
                .Where(item =>
                    (item.Accessibility == Accessibility.Public || item.Accessibility == Accessibility.Implicit)
                    && MemberTypes.Contains(item.DeclarationType))
                .OrderBy(o => o.Selection.StartLine)
                .ThenBy(t => t.Selection.StartColumn)
                .Select(d => new InterfaceMember(d, codeBuilder))
                .ToList());
        }
    }
}
