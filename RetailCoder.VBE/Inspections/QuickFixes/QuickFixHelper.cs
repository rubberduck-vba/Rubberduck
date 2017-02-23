using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace Rubberduck.Inspections.QuickFixes
{
    class QuickFixHelper
    {
        private Declaration _target;
        private QualifiedSelection _selection;
        private HashSet<string> _identifierNamesDeclaredInProcedureScope;
        private ICodeModule _module;
        public QuickFixHelper(Declaration target, QualifiedSelection selection)
        {
            _target = target;
            _selection = selection;
            _identifierNamesDeclaredInProcedureScope = new HashSet<string>();
            _module = selection.QualifiedName.Component.CodeModule;
        }
        /// <summary>
        /// Returns the CodeModule string located at lineNumber
        /// of the newly inserted line.
        /// </summary>
        public string GetModuleLine(int lineNumber)
        {
            return _module.GetLines(lineNumber, 1);
        }
        /// <summary>
        /// Inserts a string/line in a Code Module after the lineNumber provided.  Returns the line number
        /// of the newly inserted line.
        /// </summary>
        public int InsertAfterCodeModuleLine(int lineNumber, string[] newContent)
        {
            foreach( string line in newContent)
            {
                _module.InsertLines(++lineNumber, line);
            }
            return lineNumber;
        }
        /// <summary>
        /// Replaces a string/line in a Code Module and returns the line number
        /// of the replaced line.
        /// </summary>
        public int ReplaceModuleLine(int lineNumber, string newContent)
        {
            _module.ReplaceLine(lineNumber, newContent);
            return lineNumber;
        }
        /// <summary>
        /// Replaces a the TerminalNode Text value found at the location specified by the 
        /// TerminalNode context.
        /// </summary>
        public void ReplaceTerminalNodeTextInCodeModule(ITerminalNode terminalNode, string replacement)
        {
            var newCode = GenerateTerminalNodeTextReplacementLine(terminalNode, replacement);
            ReplaceModuleLine(terminalNode.Symbol.Line, newCode);
        }
        /// <summary>
        /// Replaces a the identifierReference.Name found at the location specified 
        /// by the identifierReference context.
        /// </summary>
        public void ReplaceIdentifierReferenceNameInModule(IdentifierReference identifierReference, string replacementName)
        {
            var newCode = GenerateIdentifierReferenceReplacementLine(identifierReference, replacementName);
            ReplaceModuleLine(identifierReference.Selection.StartLine, newCode);
        }
        /// <summary>
        /// Returns an array of IdentifierContext Names used within the procedure context.
        /// </summary>
        public string[] GetIdentifierNamesAccessibleToProcedureContext()
        {
            return GetVariableNamesAccessibleToProcedureContext(_target.Context.Parent.Parent).ToArray();
        }
        public IReadOnlyList<VBAParser.BlockStmtContext> GetBlockStmtContextsForContext(RuleContext context)
        {
            if (context is VBAParser.SubStmtContext)
            {
                return ((VBAParser.SubStmtContext)context).block().blockStmt();
            }
            else if (context is VBAParser.FunctionStmtContext)
            {
                return ((VBAParser.FunctionStmtContext)context).block().blockStmt();
            }
            else if (context is VBAParser.PropertyLetStmtContext)
            {
                return ((VBAParser.PropertyLetStmtContext)context).block().blockStmt();
            }
            else if (context is VBAParser.PropertyGetStmtContext)
            {
                return ((VBAParser.PropertyGetStmtContext)context).block().blockStmt();
            }
            else if (context is VBAParser.PropertySetStmtContext)
            {
                return ((VBAParser.PropertySetStmtContext)context).block().blockStmt();
            }
            return Enumerable.Empty<VBAParser.BlockStmtContext>().ToArray();
        }

        public IReadOnlyList<VBAParser.ArgContext> GetArgContextsForContext(RuleContext context)
        {
            if (context is VBAParser.SubStmtContext)
            {
                return ((VBAParser.SubStmtContext)context).argList().arg();
            }
            else if (context is VBAParser.FunctionStmtContext)
            {
                return ((VBAParser.FunctionStmtContext)context).argList().arg();
            }
            else if (context is VBAParser.PropertyLetStmtContext)
            {
                return ((VBAParser.PropertyLetStmtContext)context).argList().arg();
            }
            else if (context is VBAParser.PropertyGetStmtContext)
            {
                return ((VBAParser.PropertyGetStmtContext)context).argList().arg();
            }
            else if (context is VBAParser.PropertySetStmtContext)
            {
                return ((VBAParser.PropertySetStmtContext)context).argList().arg();
            }
            return Enumerable.Empty<VBAParser.ArgContext>().ToArray();
        }
        private string GenerateIdentifierReferenceReplacementLine(IdentifierReference identifierReference, string replacement)
        {
            var currentCode = GetModuleLine(identifierReference.Selection.StartLine);
            return ReplaceStringAtIndex(currentCode, identifierReference.IdentifierName, replacement, identifierReference.Context.Start.Column);
        }
        private string GenerateTerminalNodeTextReplacementLine(ITerminalNode terminalNode, string replacement)
        {
            var currentCode = GetModuleLine(terminalNode.Symbol.Line);
            return ReplaceStringAtIndex(currentCode, terminalNode.GetText(), replacement, terminalNode.Symbol.Column);
        }
        private string ReplaceStringAtIndex(string original, string toReplace, string replacement, int startIndex)
        {
            Debug.Assert(startIndex >= 0);
            Debug.Assert(original.Contains(toReplace));

            int stopIndex = startIndex + toReplace.Length - 1;
            var prefix = original.Substring(0, startIndex);
            var suffix = (stopIndex >= original.Length) ? string.Empty : original.Substring(stopIndex + 1);
            var toBeReplaced = original.Substring(startIndex, stopIndex - startIndex + 1);

            Debug.Assert(toBeReplaced.IndexOf(toReplace) == 0);
            return prefix + toBeReplaced.Replace(toReplace, replacement) + suffix;
        }
        private HashSet<string> GetVariableNamesAccessibleToProcedureContext(RuleContext ruleContext)
        {
            var allIdentifiers = new HashSet<string>();

            var blockStmtIdentifiers = GetBlockStmtIdentifiers(ruleContext);
            allIdentifiers.UnionWith(blockStmtIdentifiers);

            var potentiallyUnreferencedParameters = GetArgContextIdentifiers(ruleContext);
            allIdentifiers.UnionWith(potentiallyUnreferencedParameters);

            //TODO: add module and global scope variableNames

            return allIdentifiers;
        }
        private HashSet<string> GetBlockStmtIdentifiers(RuleContext ruleContext)
        {
            var blocks = GetBlockStmtContextsForContext(ruleContext);

            return GetIdentifierNames(blocks);
        }
        private HashSet<string> GetArgContextIdentifiers(RuleContext ruleContext)
        {
            var args = GetArgContextsForContext(ruleContext);

            return GetIdentifierNames(args);
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
            foreach(IParseTree child in children)
            {
                if(child is VBAParser.IdentifierContext)
                {
                    var childName = Identifier.GetName((VBAParser.IdentifierContext)child);
                    if (!tokenValues.Contains(childName))
                    {
                        results.Add(childName);
                    }
                }
                else
                {
                    if(!(child is TerminalNodeImpl))
                    {
                        results.UnionWith(GetIdentifierNames((RuleContext)child));
                    }
                }
            }
            return results;
        }
        private List<IParseTree> GetChildren(RuleContext ruleCtx)
        {
            var result = new List<IParseTree>();
            for(int index = 0; index < ruleCtx.ChildCount; index++)
            {
                result.Add(ruleCtx.GetChild(index));
            }
            return result;
        }
    }
}
