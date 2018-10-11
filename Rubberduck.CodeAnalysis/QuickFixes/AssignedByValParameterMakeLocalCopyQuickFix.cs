using System;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI.Refactorings;
using Rubberduck.Common;
using System.Windows.Forms;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using System.Diagnostics;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class AssignedByValParameterMakeLocalCopyQuickFix : QuickFixBase
    {
        private readonly IAssignedByValParameterQuickFixDialogFactory _dialogFactory;
        private readonly RubberduckParserState _parserState;
        private Declaration _quickFixTarget;

        public AssignedByValParameterMakeLocalCopyQuickFix(RubberduckParserState state, IAssignedByValParameterQuickFixDialogFactory dialogFactory)
            : base(typeof(AssignedByValParameterInspection))
        {
            _dialogFactory = dialogFactory;
            _parserState = state;
        }

        public override void Fix(IInspectionResult result)
        {
            Debug.Assert(result.Target.Context.Parent is VBAParser.ArgListContext);
            Debug.Assert(null != ((ParserRuleContext)result.Target.Context.Parent.Parent).GetChild<VBAParser.EndOfStatementContext>());

            _quickFixTarget = result.Target;

            var localIdentifier = PromptForLocalVariableName(result.Target);
            if (string.IsNullOrEmpty(localIdentifier))
            {
                return;
            }

            var rewriter = _parserState.GetRewriter(result.Target);
            ReplaceAssignedByValParameterReferences(rewriter, result.Target, localIdentifier);
            InsertLocalVariableDeclarationAndAssignment(rewriter, result.Target, localIdentifier);
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.AssignedByValParameterMakeLocalCopyQuickFix;

        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => false;
        public override bool CanFixInProject => false;

        private string PromptForLocalVariableName(Declaration target)
        {
            IAssignedByValParameterQuickFixDialog view = null;
            try
            {
                view = _dialogFactory.Create(target.IdentifierName, target.DeclarationType.ToString(), IsNameCollision);
                view.NewName = GetDefaultLocalIdentifier(target);
                view.ShowDialog();

                if (view.DialogResult == DialogResult.Cancel || !IsValidVariableName(view.NewName))
                {
                    return string.Empty;
                }

                return view.NewName;
            }
            finally
            {
                _dialogFactory.Release(view);
            }
        }

        private bool IsNameCollision(string newName)
            => _parserState.DeclarationFinder.FindNewDeclarationNameConflicts(newName, _quickFixTarget).Any();

        private string GetDefaultLocalIdentifier(Declaration target)
        {
            var newName = $"local{target.IdentifierName.CapitalizeFirstLetter()}";
            if (IsValidVariableName(newName))
            {
                return newName;
            }

            for ( var attempt = 2; attempt < 10; attempt++)
            {
                var result = newName + attempt;
                if (IsValidVariableName(result))
                {
                    return result;
                }
            }
            return newName;
        }

        private bool IsValidVariableName(string variableName)
        {
            return VariableNameValidator.IsValidName(variableName)
                && !IsNameCollision(variableName);
        }

        private void ReplaceAssignedByValParameterReferences(IModuleRewriter rewriter, Declaration target, string localIdentifier)
        {
            foreach (var identifierReference in target.References)
            {
                rewriter.Replace(identifierReference.Context, localIdentifier);
            }
        }

        private void InsertLocalVariableDeclarationAndAssignment(IModuleRewriter rewriter, Declaration target, string localIdentifier)
        {
            var localVariableDeclaration = $"{Tokens.Dim} {localIdentifier} {Tokens.As} {target.AsTypeName}";

            var requiresAssignmentUsingSet =
                target.References.Any(refItem => VariableRequiresSetAssignmentEvaluator.RequiresSetAssignment(refItem, _parserState));

            var localVariableAssignment =
                $"{(requiresAssignmentUsingSet ? $"{Tokens.Set} " : string.Empty)}{localIdentifier} = {target.IdentifierName}";

            var endOfStmtCtxt = ((ParserRuleContext)target.Context.Parent.Parent).GetChild<VBAParser.EndOfStatementContext>();
            var eosContent = endOfStmtCtxt.GetText();
            var idxLastNewLine = eosContent.LastIndexOf(Environment.NewLine, StringComparison.InvariantCultureIgnoreCase);
            var endOfStmtCtxtComment = eosContent.Substring(0, idxLastNewLine);
            var endOfStmtCtxtEndFormat = eosContent.Substring(idxLastNewLine);

            var insertCtxt = ((ParserRuleContext) target.Context.Parent.Parent).GetChild<VBAParser.AsTypeClauseContext>()
                ?? (ParserRuleContext) target.Context.Parent;

            rewriter.Remove(endOfStmtCtxt);
            rewriter.InsertAfter(insertCtxt.Stop.TokenIndex, $"{endOfStmtCtxtComment}{endOfStmtCtxtEndFormat}{localVariableDeclaration}" + $"{endOfStmtCtxtEndFormat}{localVariableAssignment}{endOfStmtCtxtEndFormat}");
        }
    }
}
