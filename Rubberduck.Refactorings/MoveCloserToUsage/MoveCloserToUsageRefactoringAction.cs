using System;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.DeleteDeclarations;

namespace Rubberduck.Refactorings.MoveCloserToUsage
{
    public class MoveCloserToUsageRefactoringAction : RefactoringActionBase<MoveCloserToUsageModel>
    {
        private readonly ICodeOnlyRefactoringAction<DeleteDeclarationsModel> _deleteDeclarationsRefactoringAction;
        public MoveCloserToUsageRefactoringAction(DeleteDeclarationsRefactoringAction deleteDeclarationsRefactoringAction, IRewritingManager rewritingManager) 
            : base(rewritingManager)
        {
            _deleteDeclarationsRefactoringAction = deleteDeclarationsRefactoringAction;
        }

        protected override void Refactor(MoveCloserToUsageModel model, IRewriteSession rewriteSession)
        {
            var variable = model.Target;
            if (!(model.DeclarationStatement == Tokens.Dim || model.DeclarationStatement == Tokens.Static))
            {
                throw new ArgumentException("Invalid value - DeclarationStatement required");
            }

            InsertNewDeclaration(variable, rewriteSession, model.DeclarationStatement);

            _deleteDeclarationsRefactoringAction.Refactor(new DeleteDeclarationsModel(variable), rewriteSession);

            UpdateQualifiedCalls(variable, rewriteSession);
        }

        private void InsertNewDeclaration(VariableDeclaration target, IRewriteSession rewriteSession, string DeclarationStatement)
        {
            var subscripts = target.Context.GetDescendent<VBAParser.BoundsListContext>()?.GetText() ?? string.Empty;
            var identifier = target.IsArray ? $"{target.IdentifierName}({subscripts})" : target.IdentifierName;

            var newVariable = target.AsTypeContext is null
                ? $"{DeclarationStatement} {identifier} {Tokens.As} {Tokens.Variant}"
                : $"{DeclarationStatement} {identifier} {Tokens.As} {(target.IsSelfAssigned ? Tokens.New + " " : string.Empty)}{target.AsTypeNameWithoutArrayDesignator}";

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

        private void UpdateQualifiedCalls(VariableDeclaration target, IRewriteSession rewriteSession)
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