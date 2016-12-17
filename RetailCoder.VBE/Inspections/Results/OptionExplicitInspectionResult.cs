using System.Collections.Generic;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Results
{
    public class OptionExplicitInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<QuickFixBase> _quickFixes; 

        public OptionExplicitInspectionResult(IInspection inspection, QualifiedModuleName qualifiedName) 
            : base(inspection, new CommentNode(string.Empty, Tokens.CommentMarker, new QualifiedSelection(qualifiedName, Selection.Home)))
        {
            _quickFixes = new[]
            {
                new OptionExplicitQuickFix(Context, QualifiedSelection), 
            };
        }

        public override IEnumerable<QuickFixBase> QuickFixes { get { return _quickFixes; } }

        public override string Description
        {
            get { return string.Format(InspectionsUI.OptionExplicitInspectionResultFormat, QualifiedName.ComponentName); }
        }
    }
}
