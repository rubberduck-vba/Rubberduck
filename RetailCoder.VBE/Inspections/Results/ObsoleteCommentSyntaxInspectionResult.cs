using System.Collections.Generic;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.Results
{
    public class ObsoleteCommentSyntaxInspectionResult : InspectionResultBase
    {
        private readonly IEnumerable<QuickFixBase> _quickFixes;

        public ObsoleteCommentSyntaxInspectionResult(IInspection inspection, CommentNode comment) 
            : base(inspection, comment)
        {
            _quickFixes = new QuickFixBase[]
            {
                new ReplaceObsoleteCommentMarkerQuickFix(Context, QualifiedSelection, comment),
                new RemoveCommentQuickFix(Context, QualifiedSelection, comment), 
                new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName), 
            };
        }

        public override IEnumerable<QuickFixBase> QuickFixes { get { return _quickFixes; } }

        public override string Description
        {
            get { return InspectionsUI.ObsoleteCommentSyntaxInspectionResultFormat; }
        }
    }
}
