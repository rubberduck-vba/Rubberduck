using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class ObsoleteGlobalInspectionResult : CodeInspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes;

        public ObsoleteGlobalInspectionResult(string inspection, CodeInspectionSeverity type, QualifiedContext<ParserRuleContext> context)
            : base(inspection, type, context.ModuleName, context.Context)
        {
            _quickFixes = new[]
            {
                new ReplaceGlobalModifierQuickFix(Context, QualifiedSelection)
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }
    }

    public class ReplaceGlobalModifierQuickFix : CodeInspectionQuickFix
    {
        public ReplaceGlobalModifierQuickFix(ParserRuleContext context, QualifiedSelection selection)
            : base(context, selection, RubberduckUI.Inspections_ChangeGlobalAccessModifierToPublic)
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