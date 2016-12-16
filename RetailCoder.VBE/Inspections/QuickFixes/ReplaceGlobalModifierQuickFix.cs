using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    public class ReplaceGlobalModifierQuickFix : QuickFixBase
    {
        public ReplaceGlobalModifierQuickFix(ParserRuleContext context, QualifiedSelection selection)
            : base(context, selection, InspectionsUI.ObsoleteGlobalInspectionQuickFix)
        {
        }

        public override void Fix()
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