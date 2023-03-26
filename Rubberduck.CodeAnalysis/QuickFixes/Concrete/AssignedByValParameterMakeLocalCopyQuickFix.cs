using System;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using Antlr4.Runtime;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.UI.Refactorings;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Introduces a new local variable and assigns it at the top of the procedure scope, then updates all parameter references to refer to the new local variable.
    /// </summary>
    /// <inspections>
    /// <inspection name="AssignedByValParameterInspection" />
    /// </inspections>
    /// <canfix multiple="false" procedure="false" module="false" project="false" all="false" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal value As Long)
    ///     Debug.Print value
    ///     value = 42
    ///     Debug.Print value
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal value As Long)
    ///     Dim localValue As Long
    ///     localValue = value
    ///     Debug.Print localValue
    ///     localValue = 42
    ///     Debug.Print localValue
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class AssignedByValParameterMakeLocalCopyQuickFix : QuickFixBase
    {
        private readonly IAssignedByValParameterQuickFixDialogFactory _dialogFactory;
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private Declaration _quickFixTarget;

        public AssignedByValParameterMakeLocalCopyQuickFix(IDeclarationFinderProvider declarationFinderProvider, IAssignedByValParameterQuickFixDialogFactory dialogFactory)
            : base(typeof(AssignedByValParameterInspection))
        {
            _dialogFactory = dialogFactory;
            _declarationFinderProvider = declarationFinderProvider;
        }

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            Debug.Assert(result.Target.Context.Parent is VBAParser.ArgListContext);
            Debug.Assert(null != ((ParserRuleContext)result.Target.Context.Parent.Parent).GetChild<VBAParser.EndOfStatementContext>());

            _quickFixTarget = result.Target;

            var localIdentifier = PromptForLocalVariableName(result.Target);
            if (string.IsNullOrEmpty(localIdentifier))
            {
                return;
            }

            var rewriter = rewriteSession.CheckOutModuleRewriter(result.Target.QualifiedModuleName);
            ReplaceAssignedByValParameterReferences(rewriter, result.Target, localIdentifier);
            InsertLocalVariableDeclarationAndAssignment(rewriter, result.Target, localIdentifier);
        }

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
            => _declarationFinderProvider.DeclarationFinder.FindNewDeclarationNameConflicts(newName, _quickFixTarget).Any();

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
            return VBAIdentifierValidator.IsValidIdentifier(variableName, DeclarationType.Variable)
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
                target.References.Any(refItem => VariableRequiresSetAssignmentEvaluator.RequiresSetAssignment(refItem, _declarationFinderProvider));

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

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.AssignedByValParameterMakeLocalCopyQuickFix;

        public override bool CanFixMultiple => false;
        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => false;
        public override bool CanFixInProject => false;
        public override bool CanFixAll => false;

    }
}
