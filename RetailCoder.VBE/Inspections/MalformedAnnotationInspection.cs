using System.Collections.Generic;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Inspections
{
    public sealed class MalformedAnnotationInspection : InspectionBase, IParseTreeInspection
    {
        public MalformedAnnotationInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Error)
        {
        }

        public override string Meta { get { return InspectionsUI.MalformedAnnotationInspectionMeta; } }
        public override string Description { get { return InspectionsUI.MalformedAnnotationInspectionResultFormat; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public ParseTreeResults ParseTreeResults { get; set; }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            if (ParseTreeResults == null)
            {
                return new InspectionResultBase[] { };
            }

            var results = new List<MalformedAnnotationInspectionResult>();

            foreach (var result in ParseTreeResults.MalformedAnnotations)
            {
                var context = (VBAParser.AnnotationContext)result.Context;

                if (context.annotationName().GetText() == AnnotationType.Ignore.ToString() ||
                    context.annotationName().GetText() == AnnotationType.Folder.ToString())
                {
                    if (context.annotationArgList() == null)
                    {
                        results.Add(new MalformedAnnotationInspectionResult(this,
                            new QualifiedContext<VBAParser.AnnotationContext>(result.ModuleName,
                                context)));
                    }
                }
            }

            return results;
        }

        public class MalformedAnnotationStatementListener : VBAParserBaseListener
        {
            private readonly IList<VBAParser.AnnotationContext> _contexts = new List<VBAParser.AnnotationContext>();
            public IEnumerable<VBAParser.AnnotationContext> Contexts { get { return _contexts; } }

            public override void ExitAnnotation(VBAParser.AnnotationContext context)
            {
                if (context.annotationName() != null)
                {
                    _contexts.Add(context);
                }
            }
        }
    }

    public class MalformedAnnotationInspectionResult : InspectionResultBase
    {
        public MalformedAnnotationInspectionResult(IInspection inspection, QualifiedContext<VBAParser.AnnotationContext> qualifiedContext)
            : base(inspection, qualifiedContext.ModuleName, qualifiedContext.Context)
        {
        }

        public override string Description
        {
            get { return string.Format(Inspection.Description, ((VBAParser.AnnotationContext)Context).annotationName().GetText()).Captialize(); }
        }
    }
}
