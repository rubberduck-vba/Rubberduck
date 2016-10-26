using System.Diagnostics;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class UntypedFunctionUsageQuickFix : CodeInspectionQuickFix
    {
        public UntypedFunctionUsageQuickFix(ParserRuleContext context, QualifiedSelection selection) 
            : base(context, selection, string.Format(InspectionsUI.QuickFixUseTypedFunction_, context.GetText(), GetNewSignature(context)))
        {
        }

        public override void Fix()
        {
            var originalInstruction = Context.GetText();
            var newInstruction = GetNewSignature(Context);
            var selection = Selection.Selection;

            var module = Selection.QualifiedName.Component.CodeModule;
            var lines = module.GetLines(selection.StartLine, selection.LineCount);

            var result = lines.Remove(Context.Start.Column, originalInstruction.Length)
                .Insert(Context.Start.Column, newInstruction);
            module.ReplaceLine(selection.StartLine, result);
        }

        private static string GetNewSignature(ParserRuleContext context)
        {
            Debug.Assert(context != null);

            return context.children.Aggregate(string.Empty, (current, member) =>
            {
                var isIdentifierNode = member is VBAParser.IdentifierContext;
                return current + member.GetText() + (isIdentifierNode ? "$" : string.Empty);
            });
        }
    }
}