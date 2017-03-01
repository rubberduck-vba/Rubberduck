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
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public class AssignedByValParameterMakeLocalCopyQuickFix : QuickFixBase
    {
        private readonly Declaration _target;
        private readonly IAssignedByValParameterQuickFixDialogFactory _dialogFactory;
        private readonly RubberduckParserState _parserState;
        private string _localCopyVariableName;
        private string[] _variableNamesAccessibleToProcedureContext;

        public AssignedByValParameterMakeLocalCopyQuickFix(Declaration target, QualifiedSelection selection, RubberduckParserState parserState, IAssignedByValParameterQuickFixDialogFactory dialogFactory)
            : base(target.Context, selection, InspectionsUI.AssignedByValParameterMakeLocalCopyQuickFix)
        {
            _target = target;
            _dialogFactory = dialogFactory;
            _parserState = parserState;
            _variableNamesAccessibleToProcedureContext = GetUserDefinedNamesAccessibleToProcedureContext(_target.Context.Parent.Parent);
            SetValidLocalCopyVariableNameSuggestion();
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

        private void SetValidLocalCopyVariableNameSuggestion()
        {
            _localCopyVariableName = "x" + _target.IdentifierName.CapitalizeFirstLetter();
            if (VariableNameIsValid(_localCopyVariableName)) { return; }

            //If the initial suggestion is not valid, keep pre-pending x's until it is
            for ( int attempt = 2; attempt < 10; attempt++) 
            {
                _localCopyVariableName = "x" + _localCopyVariableName;
                if (VariableNameIsValid(_localCopyVariableName))
                {
                    return;
                }
            }
            //if "xxFoo" to "xxxxxxxxxxFoo" isn't unique, give up and go with the original suggestion.
            //The QuickFix will leave the code as-is unless it receives a name that is free of conflicts
            _localCopyVariableName = "x" + _target.IdentifierName.CapitalizeFirstLetter();
        }

        private bool VariableNameIsValid(string variableName)
        {
            var validator = new VariableNameValidator(variableName);
            return validator.IsValidName()
                && !_variableNamesAccessibleToProcedureContext
                    .Any(name => name.Equals(variableName, System.StringComparison.InvariantCultureIgnoreCase));
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
            var blocks = QuickFixHelper.GetBlockStmtContexts(_target.Context.Parent.Parent);
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

        private string[] GetUserDefinedNamesAccessibleToProcedureContext(RuleContext ruleContext)
        {
            var allIdentifiers = new HashSet<string>();

            //Locally declared variable names
            var blocks = QuickFixHelper.GetBlockStmtContexts(ruleContext);

            var blockStmtIdentifierContexts = GetIdentifierContexts(blocks);
            var blockStmtIdentifiers = GetVariableNamesFromRuleContexts(blockStmtIdentifierContexts.ToArray());

            allIdentifiers.UnionWith(blockStmtIdentifiers);

            //The parameters of the procedure that are unreferenced in the procedure body
            var args = QuickFixHelper.GetArgContexts(ruleContext);

            var potentiallyUnreferencedIdentifierContexts = GetIdentifierContexts(args);
            var potentiallyUnreferencedParameters = GetVariableNamesFromRuleContexts(potentiallyUnreferencedIdentifierContexts.ToArray());

            allIdentifiers.UnionWith(potentiallyUnreferencedParameters);

            //All declarations within the same module, but outside of all procedures (e.g., member variables, procedure names)
            var sameModuleDeclarations = _parserState.AllUserDeclarations
                    .Where(item => item.ComponentName == _target.ComponentName
                    && !IsProceduralContext(item.ParentDeclaration.Context))
                    .ToList();

            allIdentifiers.UnionWith(sameModuleDeclarations.Select(d => d.IdentifierName));

            //Public declarations anywhere within the project other than Public members and 
            //procedures of Class  modules
            var allPublicDeclarations = _parserState.AllUserDeclarations
                .Where(item => (item.Accessibility == Accessibility.Public
                || ((item.Accessibility == Accessibility.Implicit) 
                && item.ParentScopeDeclaration is ProceduralModuleDeclaration))
                && !(item.ParentScopeDeclaration is ClassModuleDeclaration))
                .ToList();

            allIdentifiers.UnionWith(allPublicDeclarations.Select(d => d.IdentifierName));

            return allIdentifiers.ToArray();
        }

        private HashSet<string> GetVariableNamesFromRuleContexts(RuleContext[] ruleContexts)
        {
            var tokenValues = typeof(Tokens).GetFields().Select(item => item.GetValue(null)).Cast<string>().Select(item => item);
            var results = new HashSet<string>();

            foreach( var ruleContext in ruleContexts)
            {
                var name = Identifier.GetName((VBAParser.IdentifierContext)ruleContext);
                if (!tokenValues.Contains(name))
                {
                    results.Add(name);
                }
            }
            return results;
        }

        private HashSet<RuleContext> GetIdentifierContexts(IReadOnlyList<RuleContext> ruleContexts)
        {
            var identifiers = new HashSet<RuleContext>();
            foreach (RuleContext ruleContext in ruleContexts)
            {
                var identifiersForThisContext = GetIdentifierContexts(ruleContext);
                identifiers.UnionWith(identifiersForThisContext);
            }
            return identifiers;
        }

        private HashSet<RuleContext> GetIdentifierContexts(RuleContext ruleContext)
        {
            //Recursively work through the tree to get all IdentifierContexts
            var results = new HashSet<RuleContext>();
            var children = GetChildren(ruleContext);

            foreach (IParseTree child in children)
            {
                if (child is VBAParser.IdentifierContext)
                {
                    var childName = Identifier.GetName((VBAParser.IdentifierContext)child);
                    results.Add((RuleContext)child);
                }
                else
                {
                    if (!(child is TerminalNodeImpl))
                    {
                        results.UnionWith(GetIdentifierContexts((RuleContext)child));
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
        private bool IsProceduralContext(RuleContext context)
        {
            if (context is VBAParser.SubStmtContext)
            {
                return true;
            }
            else if (context is VBAParser.FunctionStmtContext)
            {
                return true;
            }
            else if (context is VBAParser.PropertyLetStmtContext)
            {
                return true;
            }
            else if (context is VBAParser.PropertyGetStmtContext)
            {
                return true;
            }
            else if (context is VBAParser.PropertySetStmtContext)
            {
                return true;
            }
            return false;
        }
    }
}
