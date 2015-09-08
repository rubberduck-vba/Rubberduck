using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Nodes;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class OptionExplicitInspectionResult : CodeInspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes; 

        public OptionExplicitInspectionResult(string inspection, CodeInspectionSeverity type, QualifiedModuleName qualifiedName) 
            : base(inspection, type, new CommentNode("", new QualifiedSelection(qualifiedName, Selection.Home)))
        {
            _quickFixes = new[]
            {
                new OptionExplicitQuickFix(Context, QualifiedSelection), 
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }
    }

    public class OptionExplicitQuickFix : CodeInspectionQuickFix
    {
        // todo: add a resource string for "Specify Option Explicit".
        public OptionExplicitQuickFix(ParserRuleContext context, QualifiedSelection selection) 
            : base(context, selection, RubberduckUI.OptionExplicit)
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
    }
}