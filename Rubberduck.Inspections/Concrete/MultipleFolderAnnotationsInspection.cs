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
    public sealed class MultipleFolderAnnotationsInspection : InspectionBase, IParseTreeInspection
    {
        public MultipleFolderAnnotationsInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Error) { }

        public override CodeInspectionType InspectionType => CodeInspectionType.MaintainabilityAndReadabilityIssues;

        public IInspectionListener Listener { get; } = new FolderAnnotationStatementListener();

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            return Listener.Contexts.GroupBy(s => s.ModuleName)
                .Where(g => g.Count() > 1)
                .Select(r => new QualifiedContextInspectionResult(this,
                    string.Format(InspectionsUI.MultipleFolderAnnotationsInspectionResultFormat, r.First().ModuleName.ComponentName),
                    State,
                    r.First()));
        }

        public class FolderAnnotationStatementListener : VBAParserBaseListener, IInspectionListener
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
                if (context.annotationName()?.GetText() == nameof(AnnotationType.Folder))
                {
                    _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
                }
            }
        }
    }
}
