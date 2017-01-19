using System.Collections.Generic;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.Results
{
    public class ObsoleteCommentSyntaxInspectionResult : InspectionResultBase
    {
        private IEnumerable<QuickFixBase> _quickFixes;
        private readonly CommentNode _comment;

        public ObsoleteCommentSyntaxInspectionResult(IInspection inspection, CommentNode comment) 
            : base(inspection, comment)
        {
            _comment = comment;
        }

        public override IEnumerable<QuickFixBase> QuickFixes
        {
            get
            {
                return _quickFixes ?? (_quickFixes = new QuickFixBase[]
                {
                    new ReplaceObsoleteCommentMarkerQuickFix(Context, QualifiedSelection, _comment),
                    new RemoveCommentQuickFix(Context, QualifiedSelection, _comment), 
                    new IgnoreOnceQuickFix(Context, QualifiedSelection, Inspection.AnnotationName)
                });
            }
        }

        public override string Description
        {
            get { return InspectionsUI.ObsoleteCommentSyntaxInspectionResultFormat; }
        }
    }
}
