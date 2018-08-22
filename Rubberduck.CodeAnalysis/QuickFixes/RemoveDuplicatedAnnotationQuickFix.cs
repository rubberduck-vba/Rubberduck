using System;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public class RemoveDuplicatedAnnotationQuickFix : QuickFixBase
    {
        private readonly RubberduckParserState _state;

        public RemoveDuplicatedAnnotationQuickFix(RubberduckParserState state)
            : base(typeof(DuplicatedAnnotationInspection))
        {
            _state = state;
        }

        public override void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.QualifiedSelection.QualifiedName);

            var duplicateAnnotations = result.Target.Annotations
                .GroupBy(annotation => annotation.AnnotationType)
                .Where(group => !group.First().AllowMultiple && group.Count() > 1)
                .SelectMany(group => group.Take(group.Count() - 1));

            foreach (var annotation in duplicateAnnotations)
            {
                // Remove also the annotation marker
                var annotationList = (VBAParser.AnnotationListContext)annotation.Context.Parent;
                var index = Array.IndexOf(annotationList.annotation(), annotation.Context);
                rewriter.Remove(annotationList.AT(index));

                rewriter.Remove(annotation.Context);
            }
        }

        public override string Description(IInspectionResult result) =>
            Resources.Inspections.QuickFixes.RemoveDuplicatedAnnotationQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}
