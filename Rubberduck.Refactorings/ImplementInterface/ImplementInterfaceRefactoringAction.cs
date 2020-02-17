using System.Linq;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.AddInterfaceImplementations;

namespace Rubberduck.Refactorings.ImplementInterface
{
    public class ImplementInterfaceRefactoringAction : RefactoringActionBase<ImplementInterfaceModel>
    {
        private readonly ICodeOnlyRefactoringAction<AddInterfaceImplementationsModel> _addImplementationsRefactoringAction;

        public ImplementInterfaceRefactoringAction(
            AddInterfaceImplementationsRefactoringAction addImplementationsRefactoringAction,
            IRewritingManager rewritingManager)
            : base(rewritingManager)
        {
            _addImplementationsRefactoringAction = addImplementationsRefactoringAction;
        }

        protected override void Refactor(ImplementInterfaceModel model, IRewriteSession rewriteSession)
        {
            ImplementMissingMembers(model.TargetInterface, model.TargetClass, rewriteSession);
        }

        private void ImplementMissingMembers(ModuleDeclaration targetInterface, ModuleDeclaration targetClass, IRewriteSession rewriteSession)
        {
            var implemented = targetClass.Members
                .Where(decl => decl is ModuleBodyElementDeclaration member && ReferenceEquals(member.InterfaceImplemented, targetInterface))
                .Cast<ModuleBodyElementDeclaration>()
                .Select(member => member.InterfaceMemberImplemented).ToList();

            var interfaceMembers = targetInterface.Members.OrderBy(member => member.Selection.StartLine)
                .ThenBy(member => member.Selection.StartColumn);

            var nonImplementedMembers = interfaceMembers.Where(member => !implemented.Contains(member));

            var addMembersModel = new AddInterfaceImplementationsModel(targetClass.QualifiedModuleName, targetInterface.IdentifierName, nonImplementedMembers.ToList());
            _addImplementationsRefactoringAction.Refactor(addMembersModel, rewriteSession);
        }
    }
}