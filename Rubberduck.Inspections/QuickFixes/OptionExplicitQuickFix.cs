using System;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    public class OptionExplicitQuickFix : IQuickFix
    {
        public OptionExplicitQuickFix(ParserRuleContext context, QualifiedSelection selection)
            : base(context, selection, InspectionsUI.OptionExplicitQuickFix)
        {
        }

        public void Fix(IInspectionResult result)
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