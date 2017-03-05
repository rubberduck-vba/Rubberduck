using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Inspections
{
    public sealed class MissingAnnotationArgumentInspection : InspectionBase, IParseTreeInspection
    {
        private IEnumerable<QualifiedContext> _parseTreeResults;
 
        public MissingAnnotationArgumentInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Error)
        {
        }

        public override string Meta { get { return InspectionsUI.MissingAnnotationArgumentInspectionMeta; } }
        public override string Description { get { return InspectionsUI.MissingAnnotationArgumentInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public IEnumerable<QualifiedContext<VBAParser.AnnotationContext>> ParseTreeResults { get { return _parseTreeResults.OfType<QualifiedContext<VBAParser.AnnotationContext>>(); } }

        public void SetResults(IEnumerable<QualifiedContext> results)
        {
            _parseTreeResults = results;
        }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            if (ParseTreeResults == null)
            {
                return new InspectionResultBase[] { };
            }

            return (from result in ParseTreeResults
                    let context = result.Context 
                    where context.annotationName().GetText() == AnnotationType.Ignore.ToString() 
                       || context.annotationName().GetText() == AnnotationType.Folder.ToString() 
                    where context.annotationArgList() == null 
                    select new MissingAnnotationArgumentInspectionResult(this, result)).ToList();
        }

        public class InvalidAnnotationStatementListener : VBAParserBaseListener
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
}
