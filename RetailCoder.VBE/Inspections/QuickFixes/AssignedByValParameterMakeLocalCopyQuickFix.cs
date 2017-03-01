using Rubberduck.Inspections.Abstract;
using System.Linq;
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
        private readonly IEnumerable<string> _forbiddenNames;
        private string _localCopyVariableName;

        public AssignedByValParameterMakeLocalCopyQuickFix(Declaration target, QualifiedSelection selection, IAssignedByValParameterQuickFixDialogFactory dialogFactory)
            : base(target.Context, selection, InspectionsUI.AssignedByValParameterMakeLocalCopyQuickFix)
        {
            _target = target;
            _dialogFactory = dialogFactory;
            _forbiddenNames = GetIdentifierNamesAccessibleToProcedureContext(target.Context.Parent.Parent);
            _localCopyVariableName = ComputeSuggestedName();
        }

        public override bool CanFixInModule { get { return false; } }
        public override bool CanFixInProject { get { return false; } }

        public override void Fix()
        {
            RequestLocalCopyVariableName();

            if (!VariableNameIsValid(_localCopyVariableName) || IsCancelled)
            {
                return;
            }

            ReplaceAssignedByValParameterReferences();

            InsertLocalVariableDeclarationAndAssignment();
        }

        private void RequestLocalCopyVariableName()
        {
            using( var view = _dialogFactory.Create(_target.IdentifierName, _target.DeclarationType.ToString(), _forbiddenNames))
            {
                view.NewName = _localCopyVariableName;
                view.ShowDialog();
                IsCancelled = view.DialogResult == DialogResult.Cancel;
                if (!IsCancelled)
                {
                    _localCopyVariableName = view.NewName;
                }
            }
        }

        private string ComputeSuggestedName()
        {
            var newName = "local" + _target.IdentifierName.CapitalizeFirstLetter();
            if (VariableNameIsValid(newName))
            {
                return newName;
            }

            for ( var attempt = 2; attempt < 10; attempt++)
            {
                var result = newName + attempt;
                if (VariableNameIsValid(result))
                {
                    return result;
                }
            }
            return newName;
        }

        private bool VariableNameIsValid(string variableName)
        {
            var validator = new VariableNameValidator(variableName);
            return validator.IsValidName()
                && !_forbiddenNames.Any(name => name.Equals(variableName, System.StringComparison.InvariantCultureIgnoreCase));
        }

        private void ReplaceAssignedByValParameterReferences()
        {
            var module = Selection.QualifiedName.Component.CodeModule;
            foreach (var identifierReference in _target.References)
            {
                module.ReplaceIdentifierReferenceName(identifierReference, _localCopyVariableName);
            }
        }

        private void InsertLocalVariableDeclarationAndAssignment()
        {
            var block = QuickFixHelper.GetBlockStmtContextsForContext(_target.Context.Parent.Parent).FirstOrDefault();
            if (block == null)
            {
                return;
            }

            string[] lines = { BuildLocalCopyDeclaration(), BuildLocalCopyAssignment() };
            var module = Selection.QualifiedName.Component.CodeModule;
            module.InsertLines(block.Start.Line, lines);
        }

        private string BuildLocalCopyDeclaration()
        {
            return Tokens.Dim + " " + _localCopyVariableName + " " + Tokens.As + " " + _target.AsTypeName;
        }

        private string BuildLocalCopyAssignment()
        {
            return (_target.AsTypeDeclaration is ClassModuleDeclaration ? Tokens.Set + " " : string.Empty) 
                + _localCopyVariableName + " = " + _target.IdentifierName;
        }

        private IEnumerable<string> GetIdentifierNamesAccessibleToProcedureContext(RuleContext ruleContext)
        {
            var allIdentifiers = new HashSet<string>();

            var blocks = QuickFixHelper.GetBlockStmtContextsForContext(ruleContext);

            var blockStmtIdentifiers = GetIdentifierNames(blocks);
            allIdentifiers.UnionWith(blockStmtIdentifiers);

            var args = QuickFixHelper.GetArgContextsForContext(ruleContext);

            var potentiallyUnreferencedParameters = GetIdentifierNames(args);
            allIdentifiers.UnionWith(potentiallyUnreferencedParameters);

            //TODO: add module and global scope variableNames to the list.

            return allIdentifiers.ToArray();
        }

        private IEnumerable<string> GetIdentifierNames(IEnumerable<RuleContext> ruleContexts)
        {
            var identifiers = new HashSet<string>();
            foreach (var identifiersForThisContext in ruleContexts.Select(GetIdentifierNames))
            {
                identifiers.UnionWith(identifiersForThisContext);
            }
            return identifiers;
        }

        private static HashSet<string> GetIdentifierNames(RuleContext ruleContext)
        {
            // note: this looks like something that's already handled somewhere else...

            //Recursively work through the tree to get all IdentifierContexts
            var results = new HashSet<string>();
            var tokenValues = typeof(Tokens).GetFields().Select(item => item.GetValue(null)).Cast<string>().Select(item => item).ToArray();
            var children = GetChildren(ruleContext);

            foreach (var child in children)
            {
                var context = child as VBAParser.IdentifierContext;
                if (context != null)
                {
                    var childName = Identifier.GetName(context);
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

        private static IEnumerable<IParseTree> GetChildren(IParseTree tree)
        {
            var result = new List<IParseTree>();
            for (var index = 0; index < tree.ChildCount; index++)
            {
                result.Add(tree.GetChild(index));
            }
            return result;
        }
    }
}
