using System;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI.Refactorings;
using Rubberduck.Common;
using System.Collections.Generic;
using System.Windows.Forms;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using System.Diagnostics;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class AssignedByValParameterMakeLocalCopyQuickFix : QuickFixBase
    {
        private readonly IAssignedByValParameterQuickFixDialogFactory _dialogFactory;
        private readonly RubberduckParserState _parserState;

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

            var forbiddenNames = _parserState.DeclarationFinder.GetDeclarationsWithIdentifiersToAvoid(result.Target).Select(n => n.IdentifierName);

            var localIdentifier = PromptForLocalVariableName(result.Target, forbiddenNames.ToList());
            if (string.IsNullOrEmpty(localIdentifier))
            {
                return;
            }

            var rewriter = _parserState.GetRewriter(result.Target);
            ReplaceAssignedByValParameterReferences(rewriter, result.Target, localIdentifier);
            InsertLocalVariableDeclarationAndAssignment(rewriter, result.Target, localIdentifier);
        }

        public override string Description(IInspectionResult result) => InspectionsUI.AssignedByValParameterMakeLocalCopyQuickFix;

        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => false;
        public override bool CanFixInProject => false;

        private string PromptForLocalVariableName(Declaration target, List<string> forbiddenNames)
        {
            IAssignedByValParameterQuickFixDialog view = null;
            try
            {
                view = _dialogFactory.Create(target.IdentifierName, target.DeclarationType.ToString(), forbiddenNames);
                view.NewName = GetDefaultLocalIdentifier(target, forbiddenNames);
                view.ShowDialog();

                if (view.DialogResult == DialogResult.Cancel || !IsValidVariableName(view.NewName, forbiddenNames))
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

        private string GetDefaultLocalIdentifier(Declaration target, List<string> forbiddenNames)
        {
            var newName = $"local{target.IdentifierName.CapitalizeFirstLetter()}";
            if (IsValidVariableName(newName, forbiddenNames))
            {
                return newName;
            }

            for ( var attempt = 2; attempt < 10; attempt++)
            {
                var result = newName + attempt;
                if (IsValidVariableName(result, forbiddenNames))
                {
                    return result;
                }
            }
            return newName;
        }

        private bool IsValidVariableName(string variableName, IEnumerable<string> forbiddenNames)
        {
            return VariableNameValidator.IsValidName(variableName)
                && !forbiddenNames.Any(name => name.Equals(variableName, StringComparison.InvariantCultureIgnoreCase));
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

            var localVariableAssignment = string.Format("{0}{1}",
                                                        requiresAssignmentUsingSet ? $"{Tokens.Set} " : string.Empty,
                                                        $"{localIdentifier} = {target.IdentifierName}");

            var endOfStmtCtxt = ((ParserRuleContext)target.Context.Parent.Parent).GetChild<VBAParser.EndOfStatementContext>();
            var eosContent = endOfStmtCtxt.GetText();
            var idxLastNewLine = eosContent.LastIndexOf(Environment.NewLine);
            var endOfStmtCtxtComment = eosContent.Substring(0, idxLastNewLine);
            var endOfStmtCtxtEndFormat = eosContent.Substring(idxLastNewLine);

            var insertCtxt = (ParserRuleContext)((ParserRuleContext)target.Context.Parent.Parent).GetChild<VBAParser.AsTypeClauseContext>();
            if (insertCtxt == null)
            {
                insertCtxt = (ParserRuleContext)target.Context.Parent;
            }

            rewriter.Remove(endOfStmtCtxt);
            rewriter.InsertAfter(insertCtxt.Stop.TokenIndex, $"{endOfStmtCtxtComment}{endOfStmtCtxtEndFormat}{localVariableDeclaration}" + $"{endOfStmtCtxtEndFormat}{localVariableAssignment}{endOfStmtCtxtEndFormat}");
        }
    }
}
