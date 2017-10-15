using System;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    public class OptionExplicitQuickFix : QuickFixBase
    {
        public OptionExplicitQuickFix(ParserRuleContext context, QualifiedSelection selection)
            : base(context, selection, InspectionsUI.OptionExplicitQuickFix)
        {
        }

        public override void Fix()
        {
            var module = Selection.QualifiedName.Component.CodeModule;
            if (module == null)
            {
                return;
            }

            module.InsertLines(1, Tokens.Option + ' ' + Tokens.Explicit + Environment.NewLine);
        }

        public override bool CanFixInModule { get { return false; } }
    }
}