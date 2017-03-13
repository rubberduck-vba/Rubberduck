using System;
using Rubberduck.Inspections.Abstract;
using System.Linq;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System.Windows.Forms;
using Rubberduck.UI.Refactorings;
using Rubberduck.Common;
using Antlr4.Runtime;
using System.Collections.Generic;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.PostProcessing;
using Rubberduck.Parsing.VBA;

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
            _forbiddenNames = GetIdentifierNamesAccessibleToProcedureContext();
        }

        public override bool CanFixInModule { get { return false; } }
        public override bool CanFixInProject { get { return false; } }

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
                rewriter.Replace(identifierReference, localIdentifier);
            }
        }

        private void InsertLocalVariableDeclarationAndAssignment(IModuleRewriter rewriter, string localIdentifier)
        { 
            var content = Tokens.Dim + " " + localIdentifier + " " + Tokens.As + " " + _target.AsTypeName + Environment.NewLine
                + (_target.AsTypeDeclaration is ClassModuleDeclaration ? Tokens.Set + " " : string.Empty)
                + localIdentifier + " = " + _target.IdentifierName;
            rewriter.Insert(content, ((VBAParser.ArgListContext)_target.Context.Parent).Stop.Line + 1);
        }

        private IEnumerable<string> GetIdentifierNamesAccessibleToProcedureContext()
        {
            return _parserState.AllUserDeclarations
                .Where(candidateDeclaration => 
                (
                        IsDeclarationInTheSameProcedure(candidateDeclaration, _target)
                    ||  IsDeclarationInTheSameModule(candidateDeclaration, _target)
                    ||  IsProjectGlobalDeclaration(candidateDeclaration, _target))
                 ).Select(declaration => declaration.IdentifierName).Distinct();
        }

        private bool IsDeclarationInTheSameProcedure(Declaration candidateDeclaration, Declaration scopingDeclaration)
        {
            return candidateDeclaration.ParentScope == scopingDeclaration.ParentScope;
        }

        private bool IsDeclarationInTheSameModule(Declaration candidateDeclaration, Declaration scopingDeclaration)
        {
            return candidateDeclaration.ComponentName == scopingDeclaration.ComponentName
                    && !IsDeclaredInMethodOrProperty(candidateDeclaration.ParentDeclaration.Context);
        }

        private bool IsProjectGlobalDeclaration(Declaration candidateDeclaration, Declaration scopingDeclaration)
        {
            return candidateDeclaration.ProjectName == scopingDeclaration.ProjectName
                && !(candidateDeclaration.ParentScopeDeclaration is ClassModuleDeclaration)
                && (candidateDeclaration.Accessibility == Accessibility.Public
                    || ((candidateDeclaration.Accessibility == Accessibility.Implicit)
                        && (candidateDeclaration.ParentScopeDeclaration is ProceduralModuleDeclaration)));
        }

        private bool IsDeclaredInMethodOrProperty(RuleContext procedureContext)
        {
            if (procedureContext is VBAParser.SubStmtContext)
            {
                return true;
            }
            else if (procedureContext is VBAParser.FunctionStmtContext)
            {
                return true;
            }
            else if (procedureContext is VBAParser.PropertyLetStmtContext)
            {
                return true;
            }
            else if (procedureContext is VBAParser.PropertyGetStmtContext)
            {
                return true;
            }
            else if (procedureContext is VBAParser.PropertySetStmtContext)
            {
                return true;
            }
            return false;
        }
    }
}
