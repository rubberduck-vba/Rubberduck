using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using System.Linq;

namespace Rubberduck.Inspections.QuickFixes
{
    /// <summary>
    /// Encapsulates a code inspection quickfix that changes a ByVal parameter into an explicit ByRef parameter.
    /// </summary>
    public class PassParameterByReferenceQuickFix : QuickFixBase
    {
        private Declaration _target;

        public PassParameterByReferenceQuickFix(Declaration target, QualifiedSelection selection)
            : base(target.Context, selection, InspectionsUI.PassParameterByReferenceQuickFix)
        {
            _target = target;
        }

        public override void Fix()
        {
            var argCtxt = GetArgContextForIdentifier(Context.Parent.Parent, _target.IdentifierName);

            var terminalNode = argCtxt.BYVAL();

            var replacementLine = GenerateByRefReplacementLine(terminalNode);

            ReplaceModuleLine(terminalNode.Symbol.Line, replacementLine);

        }
        private VBAParser.ArgContext GetArgContextForIdentifier(RuleContext context, string identifier)
        {
            var argList = GetArgListForContext(context);
            return argList.arg().SingleOrDefault(parameter =>
                    Identifier.GetName(parameter).Equals(identifier)
                    || Identifier.GetName(parameter).Equals("[" + identifier + "]"));
        }
        private string GenerateByRefReplacementLine(ITerminalNode terminalNode)
        {
            var module = Selection.QualifiedName.Component.CodeModule;
            var byValTokenLine = module.GetLines(terminalNode.Symbol.Line, 1);

            return ReplaceAtIndex(byValTokenLine, Tokens.ByVal, Tokens.ByRef, terminalNode.Symbol.Column);
        }
        private void ReplaceModuleLine(int lineNumber, string replacementLine)
        {
            var module = Selection.QualifiedName.Component.CodeModule;
            module.DeleteLines(lineNumber);
            module.InsertLines(lineNumber, replacementLine);
        }
        private string ReplaceAtIndex(string input, string toReplace, string replacement, int startIndex)
        {
            int stopIndex = startIndex + toReplace.Length;
            var prefix = input.Substring(0, startIndex);
            var suffix = input.Substring(stopIndex + 1);
            var tokenToBeReplaced = input.Substring(startIndex, stopIndex - startIndex + 1);
            return prefix + tokenToBeReplaced.Replace(toReplace, replacement) + suffix;
        }
        private VBAParser.ArgListContext GetArgListForContext(RuleContext context)
        {
            if (context is VBAParser.SubStmtContext)
            {
                return ((VBAParser.SubStmtContext)context).argList();
            }
            else if (context is VBAParser.FunctionStmtContext)
            {
                return ((VBAParser.FunctionStmtContext)context).argList();
            }
            else if (context is VBAParser.PropertyLetStmtContext)
            {
                return ((VBAParser.PropertyLetStmtContext)context).argList();
            }
            else if (context is VBAParser.PropertyGetStmtContext)
            {
                return ((VBAParser.PropertyGetStmtContext)context).argList();
            }
            else if (context is VBAParser.PropertySetStmtContext)
            {
                return ((VBAParser.PropertySetStmtContext)context).argList();
            }
            return null;
        }
    }
}