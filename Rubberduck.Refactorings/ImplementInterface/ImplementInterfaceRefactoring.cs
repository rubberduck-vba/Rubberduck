using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.Exceptions.ImplementInterface;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Refactorings.ImplementInterface
{
    public class ImplementInterfaceRefactoring : RefactoringBase
    {
        private readonly IRefactoringAction<ImplementInterfaceModel> _refactoringAction;
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public ImplementInterfaceRefactoring(
            ImplementInterfaceRefactoringAction refactoringAction,
            IDeclarationFinderProvider declarationFinderProvider, 
            ISelectionProvider selectionProvider)
        :base(selectionProvider)
        {
            _refactoringAction = refactoringAction;
            _declarationFinderProvider = declarationFinderProvider;
        }

        private static readonly IReadOnlyList<DeclarationType> ImplementingModuleTypes = new[]
        {
            DeclarationType.ClassModule,
            DeclarationType.UserForm, 
            DeclarationType.Document
        };

        public override void Refactor(QualifiedSelection target)
        {
            var targetInterface = _declarationFinderProvider.DeclarationFinder.FindInterface(target);

            if (targetInterface == null)
            {
                throw new NoImplementsStatementSelectedException(target);
            }

            var targetModule = _declarationFinderProvider.DeclarationFinder
                .ModuleDeclaration(target.QualifiedName);
            
            if (!ImplementingModuleTypes.Contains(targetModule.DeclarationType))
            {
                throw new InvalidDeclarationTypeException(targetModule);
            }

            var targetClass = targetModule as ClassModuleDeclaration;

            if (targetClass == null)
            {
                //This really should never happen. If it happens the declaration type enum value
                //and the type of the declaration are inconsistent.
                throw new InvalidTargetDeclarationException(targetModule);
            }

            var model = Model(targetInterface, targetClass);
            _refactoringAction.Refactor(model);
        }

        private static ImplementInterfaceModel Model(ClassModuleDeclaration targetInterface, ClassModuleDeclaration targetClass)
        {
            return new ImplementInterfaceModel(targetInterface, targetClass);
        }

        protected override Declaration FindTargetDeclaration(QualifiedSelection targetSelection)
        {
            throw new NotSupportedException();
        }

        public override void Refactor(Declaration target)
        {
            throw new NotSupportedException();
        }
    }
}
