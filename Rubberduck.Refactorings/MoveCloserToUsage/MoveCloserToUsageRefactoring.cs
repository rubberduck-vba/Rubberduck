using System;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.Exceptions.MoveCloserToUsage;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Refactorings.MoveCloserToUsage
{
    public class MoveCloserToUsageRefactoring : RefactoringBase
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;

        public MoveCloserToUsageRefactoring(
            IDeclarationFinderProvider declarationFinderProvider, 
            IRewritingManager rewritingManager,
            ISelectionProvider selectionProvider,
            ISelectedDeclarationProvider selectedDeclarationProvider)
        :base(rewritingManager, selectionProvider)
        {
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

            MoveCloserToUsage(target);
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

            var sameNameDeclarationsInModule = _declarationFinderProvider.DeclarationFinder
                .MatchName(target.IdentifierName)
                .Where(decl => decl.QualifiedModuleName.Equals(firstReference.QualifiedModuleName))
                .ToList();

            var sameNameVariablesInProcedure = sameNameDeclarationsInModule
                .Where(decl => decl.DeclarationType == DeclarationType.Variable
                               && decl.ParentScopeDeclaration.Equals(firstReference.ParentScoping));
            var conflictingSameNameVariablesInProcedure = sameNameVariablesInProcedure.FirstOrDefault();
            if (conflictingSameNameVariablesInProcedure != null)
            {
                throw new TargetDeclarationConflictsWithPreexistingDeclaration(target,
                    conflictingSameNameVariablesInProcedure);
            }

            if (target.QualifiedModuleName.Equals(firstReference.QualifiedModuleName))
            {
                //The variable is a module variable in the same module.
                //Since there is no local declaration of the of the same name in the procedure,
                //the identifier already refers to the declaration inside the method. 
                return;
            }

            //We know that the target is the only public variable of that name in a different standard module.
            var sameNameDeclarationWithModuleScope = sameNameDeclarationsInModule
                .Where(decl => decl.ParentScopeDeclaration.DeclarationType.HasFlag(DeclarationType.Module));
            var conflictingSameNameDeclarationWithModuleScope = sameNameDeclarationWithModuleScope.FirstOrDefault();
            if (conflictingSameNameDeclarationWithModuleScope != null)
            {
                throw new TargetDeclarationConflictsWithPreexistingDeclaration(target, conflictingSameNameDeclarationWithModuleScope);
            }
    }

        private void MoveCloserToUsage(Declaration target)
        {
            var rewriteSession = RewritingManager.CheckOutCodePaneSession();
            InsertNewDeclaration(target, rewriteSession);
            RemoveOldDeclaration(target, rewriteSession);
            UpdateQualifiedCalls(target, rewriteSession);
            if (!rewriteSession.TryRewrite())
            {
                throw new RewriteFailedException(rewriteSession);
            }
        }

        private void InsertNewDeclaration(Declaration target, IRewriteSession rewriteSession)
        {
            var subscripts = target.Context.GetDescendent<VBAParser.SubscriptsContext>()?.GetText() ?? string.Empty;
            var identifier = target.IsArray ? $"{target.IdentifierName}({subscripts})" : target.IdentifierName;

            var newVariable = target.AsTypeContext is null
                ? $"{Tokens.Dim} {identifier} {Tokens.As} {Tokens.Variant}"
                : $"{Tokens.Dim} {identifier} {Tokens.As} {(target.IsSelfAssigned ? Tokens.New + " " : string.Empty)}{target.AsTypeNameWithoutArrayDesignator}";

            var firstReference = target.References.OrderBy(r => r.Selection.StartLine).First();

            var enclosingBlockStatement = firstReference.Context.GetAncestor<VBAParser.BlockStmtContext>();
            var insertionIndex = enclosingBlockStatement.Start.TokenIndex;
            var insertCode = PaddedDeclaration(newVariable, enclosingBlockStatement);

            var rewriter = rewriteSession.CheckOutModuleRewriter(firstReference.QualifiedModuleName);
            rewriter.InsertBefore(insertionIndex, insertCode);
        }

        private string PaddedDeclaration(string declarationText, VBAParser.BlockStmtContext blockStmtContext)
        {
            if (blockStmtContext.TryGetPrecedingContext(out VBAParser.IndividualNonEOFEndOfStatementContext precedingEndOfStatement))
            {
                if (precedingEndOfStatement.COLON() != null)
                {
                    //You have been asking for it!
                    return $"{declarationText} : ";
                }

                var labelContext = blockStmtContext.statementLabelDefinition();
                if (labelContext != null)
                {
                    var labelAsSpace = new string(' ', labelContext.GetText().Length);
                    return $"{labelAsSpace}{blockStmtContext.whiteSpace()?.GetText()}{declarationText}{Environment.NewLine}";
                }

                var precedingWhitespaces = precedingEndOfStatement.whiteSpace();
                if (precedingWhitespaces != null && precedingWhitespaces.Length > 0)
                {
                    return $"{declarationText}{Environment.NewLine}{precedingWhitespaces[0]?.GetText()}";
                }

                return $"{declarationText}{Environment.NewLine}";
            }
            //This is the very first statement. In the context of this refactoring, this should not happen since we move a declaration into or inside a method.
            //We will handle this edge-case nonetheless and return the result with the proper indentation for this special case.
            if (blockStmtContext.TryGetPrecedingContext(out VBAParser.WhiteSpaceContext startingWhitespace))
            {
                return $"{declarationText}{Environment.NewLine}{startingWhitespace?.GetText()}";
            }

            return $"{declarationText}{Environment.NewLine}";
        }

        private void RemoveOldDeclaration(Declaration target, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(target.QualifiedModuleName);
            rewriter.Remove(target);
        }

        private void UpdateQualifiedCalls(Declaration target, IRewriteSession rewriteSession)
        {
            var references = target.References.ToList();
            var rewriter = rewriteSession.CheckOutModuleRewriter(references.First().QualifiedModuleName);
            foreach (var reference in references)
            {
                MakeReferenceUnqualified(target, reference, rewriter);
            }
        }

        private void MakeReferenceUnqualified(Declaration target, IdentifierReference reference, IModuleRewriter rewriter)
        {
            var memberAccessContext = reference.Context.GetAncestor<VBAParser.MemberAccessExprContext>();
            if (memberAccessContext == null)
            {
                return;
            }

            // member access might be to something unrelated to the rewritten target.
            // check we're not accidentally overwriting some other member-access who just happens to be a parent context
            if (memberAccessContext.unrestrictedIdentifier()?.GetText() != target.IdentifierName)
            {
                return;
            }
            var qualification = memberAccessContext.lExpression().GetText();
            if (!qualification.Equals(target.ComponentName, StringComparison.InvariantCultureIgnoreCase)
                && !qualification.Equals(target.ProjectName, StringComparison.InvariantCultureIgnoreCase)
                && !qualification.Equals($"{target.QualifiedModuleName.ProjectName}.{target.ComponentName}", StringComparison.InvariantCultureIgnoreCase))
            {
                return;
            }

            rewriter.Replace(memberAccessContext, reference.IdentifierName);
        }
    }
}
