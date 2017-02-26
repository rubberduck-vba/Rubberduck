using Rubberduck.Inspections.Abstract;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.VBEditor;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System.Windows.Forms;
using Rubberduck.UI.Refactorings;
using Rubberduck.Common;
using Antlr4.Runtime;
using System.Collections.Generic;
using Antlr4.Runtime.Tree;

namespace Rubberduck.Inspections.QuickFixes
{
    public class AssignedByValParameterMakeLocalCopyQuickFix : QuickFixBase
    {
        private readonly Declaration _target;
        private readonly IAssignedByValParameterQuickFixDialogFactory _dialogFactory;
        private string _localCopyVariableName;
        private string[] _variableNamesAccessibleToProcedureContext;

        public AssignedByValParameterMakeLocalCopyQuickFix(Declaration target, QualifiedSelection selection, IAssignedByValParameterQuickFixDialogFactory dialogFactory)
            : base(target.Context, selection, InspectionsUI.AssignedByValParameterMakeLocalCopyQuickFix)
        {
            _target = target;
            _dialogFactory = dialogFactory;
            _localCopyVariableName = "x" + _target.IdentifierName.CapitalizeFirstLetter();
            _variableNamesAccessibleToProcedureContext = GetVariableNamesAccessibleToProcedureContext(_target.Context.Parent.Parent);
        }

        public override bool CanFixInModule { get { return false; } }
        public override bool CanFixInProject { get { return false; } }

        public override void Fix()
        {
            RequestLocalCopyVariableName();

            if (!ProposedLocalVariableNameIsValid() || IsCancelled)
            {
                return;
            }

            ReplaceAssignedByValParameterReferences();

            InsertLocalVariableDeclarationAndAssignment();
        }

        private void RequestLocalCopyVariableName()
        {
            using( var view = _dialogFactory.Create(_target.IdentifierName, _target.DeclarationType.ToString()))
            {
                view.NewName = _localCopyVariableName;
                view.IdentifierNamesAlreadyDeclared = _variableNamesAccessibleToProcedureContext;
                view.ShowDialog();
                IsCancelled = view.DialogResult == DialogResult.Cancel;
                if (!IsCancelled)
                {
                    _localCopyVariableName = view.NewName;
                }
            }
        }

        private bool ProposedLocalVariableNameIsValid()
        {
            var validator = new VariableNameValidator(_localCopyVariableName);
            return validator.IsValidName()
                && !_variableNamesAccessibleToProcedureContext
                    .Any(name => name.ToUpper().Equals(_localCopyVariableName.ToUpper()));
        }

        private void ReplaceAssignedByValParameterReferences()
        {
            var module = Selection.QualifiedName.Component.CodeModule;
            foreach (IdentifierReference identifierReference in _target.References)
            {
                module.ReplaceIdentifierReferenceName(identifierReference, _localCopyVariableName);
            }
        }

        private void InsertLocalVariableDeclarationAndAssignment()
        {
            var blocks = QuickFixHelper.GetBlockStmtContextsForContext(_target.Context.Parent.Parent);
            string[] lines = { BuildLocalCopyDeclaration(), BuildLocalCopyAssignment() };
            var module = Selection.QualifiedName.Component.CodeModule;
            module.InsertLines(blocks.FirstOrDefault().Start.Line, lines);
        }

        private string BuildLocalCopyDeclaration()
        {
            return Tokens.Dim + " " + _localCopyVariableName + " " + Tokens.As 
                + " " + _target.AsTypeName;
        }
        private string BuildLocalCopyAssignment()
        {
            return (SymbolList.ValueTypes.Contains(_target.AsTypeName) ? string.Empty : Tokens.Set + " ") 
                + _localCopyVariableName + " = " + _target.IdentifierName;
        }

        private string[] GetVariableNamesAccessibleToProcedureContext(RuleContext ruleContext)
        {
            var allIdentifiers = new HashSet<string>();

            var blocks = QuickFixHelper.GetBlockStmtContextsForContext(ruleContext);

            var blockStmtIdentifiers = GetIdentifierNames(blocks);
            allIdentifiers.UnionWith(blockStmtIdentifiers);

            var args = QuickFixHelper.GetArgContextsForContext(ruleContext);

            var potentiallyUnreferencedParameters = GetIdentifierNames(args);
            allIdentifiers.UnionWith(potentiallyUnreferencedParameters);

            //TODO: add module and global scope variableNames.

            return allIdentifiers.ToArray();
        }

        private HashSet<string> GetIdentifierNames(IReadOnlyList<RuleContext> ruleContexts)
        {
            var identifiers = new HashSet<string>();
            foreach (RuleContext ruleContext in ruleContexts)
            {
                var identifiersForThisContext = GetIdentifierNames(ruleContext);
                identifiers.UnionWith(identifiersForThisContext);
            }
            return identifiers;
        }

        private HashSet<string> GetIdentifierNames(RuleContext ruleContext)
        {
            //Recursively work through the tree to get all IdentifierContexts
            var results = new HashSet<string>();
            var tokenValues = typeof(Tokens).GetFields().Select(item => item.GetValue(null)).Cast<string>().Select(item => item);
            var children = GetChildren(ruleContext);

            foreach (IParseTree child in children)
            {
                if (child is VBAParser.IdentifierContext)
                {
                    var childName = Identifier.GetName((VBAParser.IdentifierContext)child);
                    if (!tokenValues.Contains(childName))
                    {
                        results.Add(childName);
                    }
                }
                else
                {
                    if (!(child is TerminalNodeImpl))
                    {
                        results.UnionWith(GetIdentifierNames((RuleContext)child));
                    }
                }
            }
            return results;
        }

        private static List<IParseTree> GetChildren(RuleContext ruleCtx)
        {
            var result = new List<IParseTree>();
            for (int index = 0; index < ruleCtx.ChildCount; index++)
            {
                result.Add(ruleCtx.GetChild(index));
            }
            return result;
        }
    }
}
