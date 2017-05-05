using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class MissingAnnotationArgumentInspection : InspectionBase, IParseTreeInspection
    {
        public MissingAnnotationArgumentInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Error) { }

        public override CodeInspectionType InspectionType => CodeInspectionType.CodeQualityIssues;

        public IInspectionListener Listener { get; } =
            new InvalidAnnotationStatementListener();
        
        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            return (from result in Listener.Contexts
                    let context = (VBAParser.AnnotationContext)result.Context 
                    where context.annotationName().GetText() == AnnotationType.Ignore.ToString() 
                       || context.annotationName().GetText() == AnnotationType.Folder.ToString() 
                    where context.annotationArgList() == null 
                    select new QualifiedContextInspectionResult(this,
                                                string.Format(InspectionsUI.MissingAnnotationArgumentInspectionResultFormat,
                                                              ((VBAParser.AnnotationContext)result.Context).annotationName().GetText()),
                                                State,
                                                result));
        }

        public class InvalidAnnotationStatementListener : VBAParserBaseListener, IInspectionListener
        {
            private readonly List<QualifiedContext<ParserRuleContext>> _contexts = new List<QualifiedContext<ParserRuleContext>>();
            public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;

            public QualifiedModuleName CurrentModuleName { get; set; }

            public void ClearContexts()
            {
                _contexts.Clear();
            }

            public override void ExitAnnotation(VBAParser.AnnotationContext context)
            {
                if (context.annotationName() != null)
                {
                    _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
                }
            }
        }
    }
}
