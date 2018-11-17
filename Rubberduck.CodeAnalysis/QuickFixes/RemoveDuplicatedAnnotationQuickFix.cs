using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class RemoveDuplicatedAnnotationQuickFix : QuickFixBase
    {
        public RemoveDuplicatedAnnotationQuickFix()
            : base(typeof(DuplicatedAnnotationInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName);

            var duplicateAnnotations = result.Target.Annotations
                .Where(annotation => annotation.AnnotationType == result.Properties.AnnotationType)
                .OrderBy(annotation => annotation.Context.Start.StartIndex)
                .Skip(1)
                .ToList();

            var duplicatesPerAnnotationList = duplicateAnnotations
                .Select(annotation => (VBAParser.AnnotationListContext) annotation.Context.Parent)
                .Distinct()
                .ToDictionary(list => list, _ => 0);

            foreach (var annotation in duplicateAnnotations)
            {
                var annotationList = (VBAParser.AnnotationListContext)annotation.Context.Parent;

                RemoveAnnotationMarker(annotationList, annotation, rewriter);

                rewriter.Remove(annotation.Context);

                duplicatesPerAnnotationList[annotationList]++;
            }

            foreach (var pair in duplicatesPerAnnotationList)
            {
                if (OnlyQuoteRemainedFromAnnotationList(pair))
                {
                    rewriter.Remove(pair.Key);
                    rewriter.Remove(((VBAParser.CommentOrAnnotationContext) pair.Key.Parent).NEWLINE());
                }
            }
        }

        public override string Description(IInspectionResult result) =>
            Resources.Inspections.QuickFixes.RemoveDuplicatedAnnotationQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;

        private static void RemoveAnnotationMarker(VBAParser.AnnotationListContext annotationList,
            IAnnotation annotation, IModuleRewriter rewriter)
        {
            var index = Array.IndexOf(annotationList.annotation(), annotation.Context);
            rewriter.Remove(annotationList.AT(index));
        }

        private static bool OnlyQuoteRemainedFromAnnotationList(KeyValuePair<VBAParser.AnnotationListContext, int> pair)
        {
            return pair.Key.annotation().Length == pair.Value && pair.Key.commentBody() == null;
        }
    }
}
