using System;
using Rubberduck.Inspections.Abstract;
using System.Linq;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System.Windows.Forms;
using Rubberduck.UI.Refactorings;
using Rubberduck.Common;
using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.PostProcessing;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing;

namespace Rubberduck.Inspections.QuickFixes
{
    public class AssignedByValParameterMakeLocalCopyQuickFix : QuickFixBase
    {
        private readonly Declaration _target;
        private readonly IAssignedByValParameterQuickFixDialogFactory _dialogFactory;
        private readonly RubberduckParserState _parserState;
        private readonly IEnumerable<string> _forbiddenNames;

        public AssignedByValParameterMakeLocalCopyQuickFix(Declaration target, QualifiedSelection selection, RubberduckParserState parserState, IAssignedByValParameterQuickFixDialogFactory dialogFactory)
            : base(target.Context, selection, InspectionsUI.AssignedByValParameterMakeLocalCopyQuickFix)
        {
            _target = target;
            _dialogFactory = dialogFactory;
            _parserState = parserState;
            _forbiddenNames = parserState.DeclarationFinder.GetDeclarationsWithIdentifiersToAvoid(target).Select(n => n.IdentifierName);
        }

        public override bool CanFixInModule => false;
        public override bool CanFixInProject => false;

        public override void Fix()
        {
            var localIdentifier = PromptForLocalVariableName();
            if (string.IsNullOrEmpty(localIdentifier))
            {
                return;
            }

            var rewriter = _parserState.GetRewriter(_target);
            ReplaceAssignedByValParameterReferences(rewriter, localIdentifier);
            InsertLocalVariableDeclarationAndAssignment(rewriter, localIdentifier);
        }

        private string PromptForLocalVariableName()
        {
            using( var view = _dialogFactory.Create(_target.IdentifierName, _target.DeclarationType.ToString(), _forbiddenNames))
            {
                view.NewName = GetDefaultLocalIdentifier();
                view.ShowDialog();

                IsCancelled = view.DialogResult == DialogResult.Cancel;
                if (IsCancelled || !IsValidVariableName(view.NewName))
                {
                    return string.Empty;
                }

                return view.NewName;
            }
        }

        private string GetDefaultLocalIdentifier()
        {
            var newName = "local" + _target.IdentifierName.CapitalizeFirstLetter();
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
                && !_forbiddenNames.Any(name => name.Equals(variableName, StringComparison.InvariantCultureIgnoreCase));
        }

        private void ReplaceAssignedByValParameterReferences(IModuleRewriter rewriter, string localIdentifier)
        {
            foreach (var identifierReference in _target.References)
            {
                rewriter.Replace(identifierReference.Context, localIdentifier);
            }
        }

        private void InsertLocalVariableDeclarationAndAssignment(IModuleRewriter rewriter, string localIdentifier)
        {
            var localVariableDeclaration = $"{Environment.NewLine}{Tokens.Dim} {localIdentifier} {Tokens.As} {_target.AsTypeName}{Environment.NewLine}";
            
            var requiresAssignmentUsingSet =
                _target.References.Any(refItem => VariableRequiresSetAssignmentEvaluator.RequiresSetAssignment(refItem, _parserState.AllUserDeclarations));

            var localVariableAssignment = requiresAssignmentUsingSet ? $"Set {localIdentifier} = {_target.IdentifierName}" : $"{localIdentifier} = {_target.IdentifierName}";

            rewriter.InsertBefore(((ParserRuleContext)_target.Context.Parent).Stop.TokenIndex + 1, localVariableDeclaration + localVariableAssignment);
        }
    }
}
