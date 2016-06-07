using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public class OptionExplicitInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<CodeInspectionQuickFix> _quickFixes; 

        public OptionExplicitInspectionResult(IInspection inspection, QualifiedModuleName qualifiedName) 
            : base(inspection, new CommentNode(string.Empty, Tokens.CommentMarker, new QualifiedSelection(qualifiedName, Selection.Home)))
        {
            _quickFixes = new[]
            {
                new OptionExplicitQuickFix(Context, QualifiedSelection), 
            };
        }

        public override IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return _quickFixes; } }

        public override string Description
        {
            get { return string.Format(InspectionsUI.OptionExplicitInspectionResultFormat, QualifiedName.ComponentName); }
        }
    }

    public class OptionExplicitQuickFix : CodeInspectionQuickFix
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
