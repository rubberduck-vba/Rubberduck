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
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public class AssignedByValParameterMakeLocalCopyQuickFix : QuickFixBase
    {
        private readonly Declaration _target;
        private readonly IAssignedByValParameterQuickFixDialogFactory _dialogFactory;
//<<<<<<< HEAD
        private readonly RubberduckParserState _parserState;
//        private string[] _variableNamesAccessibleToProcedureContext;
//=======
        private readonly IEnumerable<string> _forbiddenNames;
//>>>>>>> rubberduck-vba/next
        private string _localCopyVariableName;

        public AssignedByValParameterMakeLocalCopyQuickFix(Declaration target, QualifiedSelection selection, RubberduckParserState parserState, IAssignedByValParameterQuickFixDialogFactory dialogFactory)
            : base(target.Context, selection, InspectionsUI.AssignedByValParameterMakeLocalCopyQuickFix)
        {
            _target = target;
            _dialogFactory = dialogFactory;
//<<<<<<< HEAD
            _parserState = parserState;
            //_variableNamesAccessibleToProcedureContext = GetUserDefinedNamesAccessibleToProcedureContext(_target.Context.Parent.Parent);
            //SetValidLocalCopyVariableNameSuggestion();
//=======
            _forbiddenNames = GetIdentifierNamesAccessibleToProcedureContext(target.Context.Parent.Parent);
           _localCopyVariableName = ComputeSuggestedName();
//>>>>>>> rubberduck-vba/next
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
            return VariableNameValidator.IsValidName(variableName)
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
//<<<<<<< HEAD
            //var blocks = QuickFixHelper.GetBlockStmtContexts(_target.Context.Parent.Parent);
//=======
            var block = QuickFixHelper.GetBlockStmtContexts(_target.Context.Parent.Parent).FirstOrDefault();
            if (block == null)
            {
                return;
            }

//>>>>>>> rubberduck-vba/next
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

//<<<<<<< HEAD
//        private string[] GetUserDefinedNamesAccessibleToProcedureContext(RuleContext ruleContext)
//=======
        private IEnumerable<string> GetIdentifierNamesAccessibleToProcedureContext(RuleContext ruleContext)
//>>>>>>> rubberduck-vba/next
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

//<<<<<<< HEAD
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
//=======
//        private IEnumerable<string> GetIdentifierNames(IEnumerable<RuleContext> ruleContexts)
//        {
//            var identifiers = new HashSet<string>();
//            foreach (var identifiersForThisContext in ruleContexts.Select(GetIdentifierNames))
//            {
//>>>>>>> rubberduck-vba/next
                identifiers.UnionWith(identifiersForThisContext);
            }
            return identifiers;
        }

//<<<<<<< HEAD
        private HashSet<RuleContext> GetIdentifierContexts(RuleContext ruleContext)
//=======
//        private static HashSet<string> GetIdentifierNames(RuleContext ruleContext)
//>>>>>>> rubberduck-vba/next
        {
            // note: this looks like something that's already handled somewhere else...

            //Recursively work through the tree to get all IdentifierContexts
//<<<<<<< HEAD
            var results = new HashSet<RuleContext>();
//=======
 //           var results = new HashSet<string>();
//            var tokenValues = typeof(Tokens).GetFields().Select(item => item.GetValue(null)).Cast<string>().Select(item => item).ToArray();
//>>>>>>> rubberduck-vba/next
            var children = GetChildren(ruleContext);

            foreach (var child in children)
            {
                var context = child as VBAParser.IdentifierContext;
                if (context != null)
                {
//<<<<<<< HEAD
                    var childName = Identifier.GetName((VBAParser.IdentifierContext)child);
                    results.Add((RuleContext)child);
//=======
 //                   var childName = Identifier.GetName(context);
//                    if (!tokenValues.Contains(childName))
//                    {
//                        results.Add(childName);
//                    }
//>>>>>>> rubberduck-vba/next
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

        private static IEnumerable<IParseTree> GetChildren(IParseTree tree)
        {
            var result = new List<IParseTree>();
            for (var index = 0; index < tree.ChildCount; index++)
            {
                result.Add(tree.GetChild(index));
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
