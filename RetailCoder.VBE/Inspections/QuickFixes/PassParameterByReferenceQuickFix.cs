using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;
using System.Text.RegularExpressions;

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
            string pattern = "^\\s*" + Tokens.ByVal + "(\\s+)";
            string rgxReplacement = Tokens.ByRef + "$1";
            Regex rgx = new Regex(pattern);

            var parameter = Context.GetText();
            var newParameter = rgx.Replace(parameter, rgxReplacement);

            var selection = Selection.Selection;

            var module = Selection.QualifiedName.Component.CodeModule;
            {
                var lines = module.GetLines(selection.StartLine, selection.LineCount);
                var result = lines.Replace(parameter, newParameter);
                module.ReplaceLine(selection.StartLine, result);
            }
        }
    }
}