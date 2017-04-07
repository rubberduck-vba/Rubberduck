using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class MissingAnnotationArgumentInspection : InspectionBase, IParseTreeInspection
    {
        private IEnumerable<QualifiedContext> _parseTreeResults;
 
        public MissingAnnotationArgumentInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Error) { }

        public override CodeInspectionType InspectionType => CodeInspectionType.CodeQualityIssues;

        public void SetResults(IEnumerable<QualifiedContext> results)
        {
            _parseTreeResults = results;
        }

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            if (_parseTreeResults == null)
            {
                return new InspectionResultBase[] { };
            }

            return (from result in _parseTreeResults.Cast<QualifiedContext<VBAParser.AnnotationContext>>()
                    let context = result.Context 
                    where context.annotationName().GetText() == AnnotationType.Ignore.ToString() 
                       || context.annotationName().GetText() == AnnotationType.Folder.ToString() 
                    where context.annotationArgList() == null 
                    select new MissingAnnotationArgumentInspectionResult(this, result)).ToList();
        }

        public class InvalidAnnotationStatementListener : VBAParserBaseListener
        {
            private readonly IList<VBAParser.AnnotationContext> _contexts = new List<VBAParser.AnnotationContext>();
            public IEnumerable<VBAParser.AnnotationContext> Contexts => _contexts;

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
