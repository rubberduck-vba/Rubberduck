using System;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.MoveCloserToUsage
{
    public class MoveCloserToUsageRefactoringAction : RefactoringActionBase<MoveCloserToUsageModel>
    {
        public MoveCloserToUsageRefactoringAction(IRewritingManager rewritingManager) 
            : base(rewritingManager)
        {}

        protected override void Refactor(MoveCloserToUsageModel model, IRewriteSession rewriteSession)
        {
            var target = model.Target;
            InsertNewDeclaration(target, rewriteSession);
            RemoveOldDeclaration(target, rewriteSession);
            UpdateQualifiedCalls(target, rewriteSession);
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