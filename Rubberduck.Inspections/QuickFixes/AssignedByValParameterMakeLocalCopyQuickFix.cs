using System;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI.Refactorings;
using Rubberduck.Common;
using System.Collections.Generic;
using System.Windows.Forms;
using Antlr4.Runtime;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.PostProcessing;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class AssignedByValParameterMakeLocalCopyQuickFix : IQuickFix
    {
        private readonly IAssignedByValParameterQuickFixDialogFactory _dialogFactory;
        private readonly RubberduckParserState _parserState;
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type> { typeof(AssignedByValParameterInspection) };

        public AssignedByValParameterMakeLocalCopyQuickFix(RubberduckParserState parserState, IAssignedByValParameterQuickFixDialogFactory dialogFactory)
        {
            _dialogFactory = dialogFactory;
            _parserState = parserState;
        }

        public IReadOnlyCollection<Type> SupportedInspections => _supportedInspections.ToList();

        public void Fix(IInspectionResult result)
        {
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

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.AssignedByValParameterMakeLocalCopyQuickFix;
        }

        public bool CanFixInProcedure { get; } = false;
        public bool CanFixInModule { get; } = false;
        public bool CanFixInProject { get; } = false;

        private string PromptForLocalVariableName(Declaration target, List<string> forbiddenNames)
        {
            using (var view = _dialogFactory.Create(target.IdentifierName, target.DeclarationType.ToString(), forbiddenNames))
            {
                view.NewName = GetDefaultLocalIdentifier(target, forbiddenNames);
                view.ShowDialog();
                
                if (view.DialogResult == DialogResult.Cancel || !IsValidVariableName(view.NewName, forbiddenNames))
                {
                    return string.Empty;
                }

                return view.NewName;
            }
        }

        private string GetDefaultLocalIdentifier(Declaration target, List<string> forbiddenNames)
        {
            var newName = "local" + target.IdentifierName.CapitalizeFirstLetter();
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
            var localVariableDeclaration = $"{Environment.NewLine}{Tokens.Dim} {localIdentifier} {Tokens.As} {target.AsTypeName}{Environment.NewLine}";
            
            var requiresAssignmentUsingSet =
                target.References.Any(refItem => VariableRequiresSetAssignmentEvaluator.RequiresSetAssignment(refItem, _parserState));

            var localVariableAssignment = requiresAssignmentUsingSet ? $"Set {localIdentifier} = {target.IdentifierName}" : $"{localIdentifier} = {target.IdentifierName}";

            rewriter.InsertBefore(((ParserRuleContext)target.Context.Parent).Stop.TokenIndex + 1, localVariableDeclaration + localVariableAssignment);
        }
    }
}
