using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;
using System.Text.RegularExpressions;
using static Rubberduck.Parsing.Grammar.VBAParser;

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
            var byValParameter = Context.GetText();

            var byRefParameter = BuildByRefParameter(byValParameter);

            ReplaceByValParameterInModule(byValParameter, byRefParameter);
        }

        private string BuildByRefParameter(string originalParameter)
        {
            var everythingAfterTheByValToken = originalParameter.Substring(Tokens.ByVal.Length);
            return Tokens.ByRef + everythingAfterTheByValToken;
        }
        private void ReplaceByValParameterInModule( string byValParameter, string byRefParameter)
        {
            var selection = Selection.Selection;
            var module = Selection.QualifiedName.Component.CodeModule;

            var lines = module.GetLines(selection.StartLine, selection.LineCount);
            var result = lines.Replace(byValParameter, byRefParameter);
            module.ReplaceLine(selection.StartLine, result);
        }
    }
}