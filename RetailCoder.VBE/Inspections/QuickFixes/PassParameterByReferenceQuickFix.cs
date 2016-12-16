using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    /// <summary>
    /// Encapsulates a code inspection quickfix that changes a ByVal parameter into an explicit ByRef parameter.
    /// </summary>
    public class PassParameterByReferenceQuickFix : QuickFixBase
    {
        public PassParameterByReferenceQuickFix(ParserRuleContext context, QualifiedSelection selection)
            : base(context, selection, InspectionsUI.PassParameterByReferenceQuickFix)
        {
        }

        public override void Fix()
        {
            var parameter = Context.GetText();
            var newContent = string.Concat(Tokens.ByRef, " ", parameter.Replace(Tokens.ByVal, string.Empty).Trim());
            var selection = Selection.Selection;

            var module = Selection.QualifiedName.Component.CodeModule;
            {
                var lines = module.GetLines(selection.StartLine, selection.LineCount);
                var result = lines.Replace(parameter, newContent);
                module.ReplaceLine(selection.StartLine, result);
            }
        }
    }
}