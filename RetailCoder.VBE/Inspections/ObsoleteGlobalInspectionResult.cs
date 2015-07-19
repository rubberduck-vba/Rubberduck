using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.UI;

namespace Rubberduck.Inspections
{
    public class ObsoleteGlobalInspectionResult : CodeInspectionResultBase
    {
        public ObsoleteGlobalInspectionResult(string inspection, CodeInspectionSeverity type, QualifiedContext<ParserRuleContext> context)
            : base(inspection, type, context.ModuleName, context.Context)
        {
        }

        public override IDictionary<string, Action> GetQuickFixes()
        {
            return new Dictionary<string, Action>
            {
                {RubberduckUI.Inspections_ChangeGlobalAccessModifierToPublic, ChangeAccessModifier}
            };
        }

        private void ChangeAccessModifier()
        {
            var module = QualifiedName.Component.CodeModule;
            if (module == null)
            {
                return;
            }

            var selection = Context.GetSelection();

            module.ReplaceLine(selection.StartLine, Tokens.Public + ' ' + Context.GetText());
        }
    }
}