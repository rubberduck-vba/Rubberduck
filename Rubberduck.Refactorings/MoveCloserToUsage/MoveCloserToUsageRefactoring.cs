using System.Linq;
using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.Exceptions.MoveCloserToUsage;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Utility;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing;
using System;

namespace Rubberduck.Refactorings.MoveCloserToUsage
{
    public class MoveCloserToUsageRefactoring : InteractiveRefactoringBase<MoveCloserToUsageModel> 
    {
        private readonly IRefactoringAction<MoveCloserToUsageModel> _refactoringAction;
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;

        public MoveCloserToUsageRefactoring(
            MoveCloserToUsageRefactoringAction refactoringAction,
            IDeclarationFinderProvider declarationFinderProvider,
            ISelectionProvider selectionProvider,
            ISelectedDeclarationProvider selectedDeclarationProvider,
            RefactoringUserInteraction<IMoveCloserToUsagePresenter, MoveCloserToUsageModel> userInteraction)
        :base(selectionProvider, userInteraction)
        {
            _refactoringAction = refactoringAction;
            _declarationFinderProvider = declarationFinderProvider;
            _selectedDeclarationProvider = selectedDeclarationProvider;
        }

        protected override Declaration FindTargetDeclaration(QualifiedSelection targetSelection)
        {
            var selectedDeclaration = _selectedDeclarationProvider.SelectedDeclaration(targetSelection);
            if (selectedDeclaration == null
                || (selectedDeclaration.DeclarationType != DeclarationType.Variable
                    && selectedDeclaration.DeclarationType != DeclarationType.Constant))
            {
                return null;
            }

            return selectedDeclaration;
        }

        public override void Refactor(Declaration target)
        {
            CheckThatTargetIsValid(target);

            var model = InitializeModel(target);
            if (DeclarationIsModuleVariableWhichRefersToDeclarationInMethod(target))
            {
                // Ask User for new Declaration Statement
                Refactor(model);
            }
            else
            {
                // Direct Refactoring
                RefactorImpl(model);
            }            
        }

        protected override MoveCloserToUsageModel InitializeModel(Declaration target)
        {
            if (!(target is VariableDeclaration variableDeclaration))
            {
                throw new ArgumentException("Invalid type - VariableDeclaration required");
            }

            var model = new MoveCloserToUsageModel(variableDeclaration, GetDefaultDeclarationStatement(target));
            return model;
        }

        protected override void RefactorImpl(MoveCloserToUsageModel model)
        {
            _refactoringAction.Refactor(model);
        }
        
        private void CheckThatTargetIsValid(Declaration target)
        {
            if (target == null)
            {
                throw new TargetDeclarationIsNullException();
            }

            if (!target.IsUserDefined)
            {
                throw new TargetDeclarationNotUserDefinedException(target);
            }

            if (target.DeclarationType != DeclarationType.Variable)
            {
                throw new InvalidDeclarationTypeException(target);
            }

            if (!target.References.Any())
            {
                throw new TargetDeclarationNotUsedException(target);
            }

            if (TargetIsReferencedFromMultipleMethods(target))
            {
                throw new TargetDeclarationUsedInMultipleMethodsException(target);
            }

            if (TargetIsInDifferentProject(target))
            {
                throw new TargetDeclarationInDifferentProjectThanUses(target);
            }

            if (TargetIsInDifferentNonStandardModule(target))
            {
                throw new TargetDeclarationInDifferentNonStandardModuleException(target);
            }

            if (TargetIsNonPrivateInNonStandardModule(target))
            {
                throw new TargetDeclarationNonPrivateInNonStandardModule(target);
            }

            CheckThatThereIsNoOtherSameNameDeclarationInScopeInReferencingMethod(target);
        }

        private static bool TargetIsReferencedFromMultipleMethods(Declaration target)
        {
            var firstReference = target.References.FirstOrDefault();

            return firstReference != null && target.References.Any(r => !Equals(r.ParentScoping, firstReference.ParentScoping));
        }

        private static bool TargetIsInDifferentProject(Declaration target)
        {
            var firstReference = target.References.FirstOrDefault();
            if (firstReference == null)
            {
                return false;
            }

            return firstReference.QualifiedModuleName.ProjectId != target.ProjectId;
        }

        private static bool TargetIsInDifferentNonStandardModule(Declaration target)
        {
            var firstReference = target.References.FirstOrDefault();
            if (firstReference == null)
            {
                return false;
            }

            return !target.QualifiedModuleName.Equals(firstReference.QualifiedModuleName)
                   && Declaration.GetModuleParent(target).DeclarationType != DeclarationType.ProceduralModule;
        }

        private static bool TargetIsNonPrivateInNonStandardModule(Declaration target)
        {
            if (!target.ParentScopeDeclaration.DeclarationType.HasFlag(DeclarationType.Module))
            {
                //local variable
                return false;
            }

            return target.Accessibility != Accessibility.Private
                && Declaration.GetModuleParent(target).DeclarationType != DeclarationType.ProceduralModule;
        }


        private void CheckThatThereIsNoOtherSameNameDeclarationInScopeInReferencingMethod(Declaration target)
        {
            var firstReference = target.References.FirstOrDefault();
            if (firstReference == null)
            {
                return;
            }

            if (firstReference.ParentScoping.Equals(target.ParentScopeDeclaration))
            {
                //The variable is already in the same scope and consequently the identifier already refers to the declaration there.
                return;
            }

            var conflictingSameNameVariablesInProcedure = GetSameNameVariablesInProcedure(target).FirstOrDefault();
            if (conflictingSameNameVariablesInProcedure != null)
            {
                throw new TargetDeclarationConflictsWithPreexistingDeclaration(target,
                    conflictingSameNameVariablesInProcedure);
            }

            if (target.QualifiedModuleName.Equals(firstReference.QualifiedModuleName))
            {
                //The variable is a module variable in the same module.
                //Since there is no local declaration with the same name in the procedure,
                //the identifier already refers to the declaration inside the method. 
                return;
            }

            //We know that the target is the only public variable of that name in a different standard module.
            var sameNameDeclarationWithModuleScope = GetSameNameDeclarationsInModule(target)
                .Where(decl => decl.ParentScopeDeclaration.DeclarationType.HasFlag(DeclarationType.Module));
            var conflictingSameNameDeclarationWithModuleScope = sameNameDeclarationWithModuleScope.FirstOrDefault();
            if (conflictingSameNameDeclarationWithModuleScope != null)
            {
                throw new TargetDeclarationConflictsWithPreexistingDeclaration(target, conflictingSameNameDeclarationWithModuleScope);
            }
        }

        private bool DeclarationIsModuleVariableWhichRefersToDeclarationInMethod(Declaration target)
        {
            var firstReference = target.References.FirstOrDefault();
            if (firstReference == null)
            {
                return false;
            }

            if (firstReference.ParentScoping.Equals(target.ParentScopeDeclaration))
            {
                //The variable is already in the same scope and consequently the identifier already refers to the declaration there.
                return false;
            }

            if (GetSameNameVariablesInProcedure(target).FirstOrDefault() != null)
            {
                return false;
            }

            if (target.QualifiedModuleName.Equals(firstReference.QualifiedModuleName))
            {
                //The variable is a module variable in the same module.
                //Since there is no local declaration with the same name in the procedure,
                //the identifier already refers to the declaration inside the method. 
                return true;
            }

            return false;
        }

        private IEnumerable<Declaration> GetSameNameDeclarationsInModule(Declaration target)
        {
            var firstReference = target.References.FirstOrDefault();
            if (firstReference == null)
            {
                return Enumerable.Empty<Declaration>();
            }

            var sameNameDeclarationsInModule = _declarationFinderProvider.DeclarationFinder
                .MatchName(target.IdentifierName)
                .Where(decl => decl.QualifiedModuleName.Equals(firstReference.QualifiedModuleName));

            return sameNameDeclarationsInModule;
        }

        private IEnumerable<Declaration> GetSameNameVariablesInProcedure(Declaration target)
        {
            var firstReference = target.References.FirstOrDefault();
            if (firstReference == null)
            {
                return Enumerable.Empty<Declaration>();
            }

            var sameNameVariablesInProcedure = GetSameNameDeclarationsInModule(target)
                .Where(decl => decl.DeclarationType == DeclarationType.Variable
                               && decl.ParentScopeDeclaration.Equals(firstReference.ParentScoping));

            return sameNameVariablesInProcedure;
        }

        static private string GetDefaultDeclarationStatement(Declaration target)
        {
            if (target.ParentDeclaration is ModuleDeclaration || IsStatic(target))
            {
                return Tokens.Static;
            }

            return Tokens.Dim;
        }

        //TODO: Add IsStatic member to VariableDeclaration - this is a copy from Rubberduck.Refractorings.AssignmentNotUsedInspection
        private static bool IsStatic(Declaration declaration)
        {
            var ctxt = declaration.Context.GetAncestor<VBAParser.VariableStmtContext>();
            if (ctxt?.STATIC() != null)
            {
                return true;
            }

            switch (declaration.ParentDeclaration.Context)
            {
                case VBAParser.FunctionStmtContext func:
                    return func.STATIC() != null;
                case VBAParser.SubStmtContext sub:
                    return sub.STATIC() != null;
                case VBAParser.PropertyLetStmtContext let:
                    return let.STATIC() != null;
                case VBAParser.PropertySetStmtContext set:
                    return set.STATIC() != null;
                case VBAParser.PropertyGetStmtContext get:
                    return get.STATIC() != null;
                default:
                    return false;
            }
        }

    }
}
