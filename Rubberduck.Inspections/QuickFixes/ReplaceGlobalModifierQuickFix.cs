using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    public class ReplaceGlobalModifierQuickFix : IQuickFix
    {
        public ReplaceGlobalModifierQuickFix(ParserRuleContext context, QualifiedSelection selection)
            : base(context, selection, InspectionsUI.ObsoleteGlobalInspectionQuickFix)
        {
        }

        public void Fix(IInspectionResult result)
        {
            var module = Selection.QualifiedName.Component.CodeModule;
            if (module == null)
            {
                return;
            }

            var selection = Context.GetSelection();

            // bug: this should make a test fail somewhere - what if identifier is one of many declarations on a line?
            module.ReplaceLine(selection.StartLine, Tokens.Public + ' ' + Context.GetText());
        }
    }
}