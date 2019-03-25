using System;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Interaction;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Refactorings.MoveCloserToUsage
{
    public class MoveCloserToUsageRefactoring : IRefactoring
    {
        private readonly ISelectionService _selectionService;
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IRewritingManager _rewritingManager;
        private readonly IMessageBox _messageBox;
        private Declaration _target;

        public MoveCloserToUsageRefactoring(IDeclarationFinderProvider declarationFinderProvider, IMessageBox messageBox, IRewritingManager rewritingManager, ISelectionService selectionService)
        {
            _selectionService = selectionService;
            _declarationFinderProvider = declarationFinderProvider;
            _rewritingManager = rewritingManager;
            _messageBox = messageBox;
        }

        public void Refactor()
        {
            var activeSelection = _selectionService.ActiveSelection();
            if (!activeSelection.HasValue)
            {
                _messageBox.NotifyWarn(RubberduckUI.MoveCloserToUsage_InvalidSelection, RubberduckUI.MoveCloserToUsage_Caption);
                return;
            }

            Refactor(activeSelection.Value);
        }

        public void Refactor(QualifiedSelection selection)
        {
            var target = _declarationFinderProvider.DeclarationFinder
                .UserDeclarations(DeclarationType.Variable)
                .FindVariable(selection);

            if (target == null)
            {
                _messageBox.NotifyWarn(RubberduckUI.MoveCloserToUsage_InvalidSelection, RubberduckUI.IntroduceParameter_Caption);
                return;
            }

            Refactor(target);
        }

        public void Refactor(Declaration target)
        {
            if (target.DeclarationType != DeclarationType.Variable)
            {
                _messageBox.NotifyWarn(RubberduckUI.MoveCloserToUsage_InvalidSelection, RubberduckUI.IntroduceParameter_Caption);
                return;
            }

            if (!target.IsUserDefined)
            {
                _messageBox.NotifyWarn(RubberduckUI.MoveCloserToUsage_TargetIsNotUserDefined, RubberduckUI.IntroduceParameter_Caption);
                return;
            }

            _target = target;
            MoveCloserToUsage();
        }

        private bool TargetIsReferencedFromMultipleMethods(Declaration target)
        {
            var firstReference = target.References.FirstOrDefault();

            return firstReference != null && target.References.Any(r => !Equals(r.ParentScoping, firstReference.ParentScoping));
        }

        private bool TargetIsInDifferentProject(Declaration target)
        {
            var firstReference = target.References.FirstOrDefault();
            if (firstReference == null)
            {
                return false;
            }

            return firstReference.QualifiedModuleName.ProjectId != target.ProjectId;
        }

        private bool TargetIsInDifferentNonStandardModule(Declaration target)
        {
            var firstReference = target.References.FirstOrDefault();
            if (firstReference == null)
            {
                return false;
            }

            return !target.QualifiedModuleName.Equals(firstReference.QualifiedModuleName)
                   && Declaration.GetModuleParent(target).DeclarationType != DeclarationType.ProceduralModule;
        }

        private bool TargetIsNonPrivateInNonStandardModule(Declaration target)
        {
            if (!_target.ParentScopeDeclaration.DeclarationType.HasFlag(DeclarationType.Module))
            {
                //local variable
                return false;
            }

            return target.Accessibility != Accessibility.Private
                && Declaration.GetModuleParent(target).DeclarationType != DeclarationType.ProceduralModule;
        }

        private bool ReferencingMethodHasOtherSameNameDeclarationInScope(Declaration target)
        {
            var firstReference = target.References.FirstOrDefault();
            if (firstReference == null)
            {
                return false;
            }

            if (firstReference.ParentScoping.Equals(_target.ParentScopeDeclaration))
            {
                //The variable is already in the same scope and consequently the identifier already refers to the declaration there.
                return false;
            }

            var sameNameDeclarationsInModule = _declarationFinderProvider.DeclarationFinder
                .MatchName(_target.IdentifierName)
                .Where(decl => decl.QualifiedModuleName.Equals(firstReference.QualifiedModuleName))
                .ToList();

            var sameNameVariablesInProcedure = sameNameDeclarationsInModule
                .Where(decl => decl.DeclarationType == DeclarationType.Variable 
                               && decl.ParentScopeDeclaration.Equals(firstReference.ParentScoping));
            if (sameNameVariablesInProcedure.Any())
            {
                return true;
            }

            if (_target.QualifiedModuleName.Equals(firstReference.QualifiedModuleName))
            {
                //The variable is a module variable in the same module.
                //Since there is no local declaration of the of the same name in the procedure,
                //the identifier already refers to the declaration inside the method. 
                return false;
            }

            //We know that the target is the only public variable of that name in a different standard module.
            var sameNameDeclarationWithModuleScope = sameNameDeclarationsInModule
                .Where(decl => decl.ParentScopeDeclaration.DeclarationType.HasFlag(DeclarationType.Module));
            return sameNameDeclarationWithModuleScope.Any();
        }

        private void MoveCloserToUsage()
        {
            if (!_target.References.Any())
            {
                var message = string.Format(RubberduckUI.MoveCloserToUsage_TargetHasNoReferences, _target.IdentifierName);
                _messageBox.NotifyWarn(message, RubberduckUI.MoveCloserToUsage_Caption);
                return;
            }

            if (TargetIsReferencedFromMultipleMethods(_target))
            {
                var message = string.Format(RubberduckUI.MoveCloserToUsage_TargetIsUsedInMultipleMethods,
                    _target.IdentifierName);
                _messageBox.NotifyWarn(message, RubberduckUI.MoveCloserToUsage_Caption);

                return;
            }

            if (TargetIsInDifferentProject(_target))
            {
                var message = string.Format(RubberduckUI.MoveCloserToUsage_TargetIsInDifferentProject,
                    _target.IdentifierName);
                _messageBox.NotifyWarn(message, RubberduckUI.MoveCloserToUsage_Caption);

                return;
            }

            if (TargetIsInDifferentNonStandardModule(_target))
            {
                var message = string.Format(RubberduckUI.MoveCloserToUsage_TargetIsInOtherNonStandardModule,
                    _target.IdentifierName);
                _messageBox.NotifyWarn(message, RubberduckUI.MoveCloserToUsage_Caption);

                return;
            }

            if (TargetIsNonPrivateInNonStandardModule(_target))
            {
                var message = string.Format(RubberduckUI.MoveCloserToUsage_TargetIsNonPrivateInNonStandardModule,
                    _target.IdentifierName);
                _messageBox.NotifyWarn(message, RubberduckUI.MoveCloserToUsage_Caption);

                return;
            }

            if (ReferencingMethodHasOtherSameNameDeclarationInScope(_target))
            {
                var message = string.Format(RubberduckUI.MoveCloserToUsage_ReferencingMethodHasSameNameDeclarationInScope,
                    _target.IdentifierName);
                _messageBox.NotifyWarn(message, RubberduckUI.MoveCloserToUsage_Caption);

                return;
            }

            var rewriteSession = _rewritingManager.CheckOutCodePaneSession();
            InsertNewDeclaration(rewriteSession);
            RemoveOldDeclaration(rewriteSession);
            UpdateQualifiedCalls(rewriteSession);
            rewriteSession.TryRewrite();
        }

        private void InsertNewDeclaration(IRewriteSession rewriteSession)
        {
            var subscripts = _target.Context.GetDescendent<VBAParser.SubscriptsContext>()?.GetText() ?? string.Empty;
            var identifier = _target.IsArray ? $"{_target.IdentifierName}({subscripts})" : _target.IdentifierName;

            var newVariable = _target.AsTypeContext is null
                ? $"{Tokens.Dim} {identifier} {Tokens.As} {Tokens.Variant}"
                : $"{Tokens.Dim} {identifier} {Tokens.As} {(_target.IsSelfAssigned ? Tokens.New + " " : string.Empty)}{_target.AsTypeNameWithoutArrayDesignator}";

            var firstReference = _target.References.OrderBy(r => r.Selection.StartLine).First();

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

        private void RemoveOldDeclaration(IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(_target.QualifiedModuleName);
            rewriter.Remove(_target);
        }

        private void UpdateQualifiedCalls(IRewriteSession rewriteSession)
        {
            var references = _target.References.ToList();
            var rewriter = rewriteSession.CheckOutModuleRewriter(references.First().QualifiedModuleName);
            foreach (var reference in references)
            {
                MakeReferenceUnqualified(reference, rewriter);
            }
        }

        private void MakeReferenceUnqualified(IdentifierReference reference, IModuleRewriter rewriter)
        {
            var memberAccessContext = reference.Context.GetAncestor<VBAParser.MemberAccessExprContext>();
            if (memberAccessContext == null)
            {
                return;
            }

            // member access might be to something unrelated to the rewritten target.
            // check we're not accidentally overwriting some other member-access who just happens to be a parent context
            if (memberAccessContext.unrestrictedIdentifier()?.GetText() != _target.IdentifierName)
            {
                return;
            }
            var qualification = memberAccessContext.lExpression().GetText();
            if (!qualification.Equals(_target.ComponentName, StringComparison.InvariantCultureIgnoreCase)
                && !qualification.Equals(_target.ProjectName, StringComparison.InvariantCultureIgnoreCase)
                && !qualification.Equals($"{_target.QualifiedModuleName.ProjectName}.{_target.ComponentName}", StringComparison.InvariantCultureIgnoreCase))
            {
                return;
            }

            rewriter.Replace(memberAccessContext, reference.IdentifierName);
        }
    }
}
