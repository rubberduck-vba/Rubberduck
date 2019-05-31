using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Warns about a malformed Rubberduck annotation that is missing an argument.
    /// </summary>
    /// <why>
    /// Some annotations require arguments; if the argument isn't specified, the annotation is nothing more than an obscure comment.
    /// </why>
    /// <example>
    /// <![CDATA[
    /// '@Folder
    /// '@ModuleDescription
    /// Option Explicit
    /// ' ...
    /// ]]>
    /// </example>
    /// <example>
    /// <![CDATA[
    /// '@Folder("MyProject.XYZ")
    /// '@ModuleDescription("This module does XYZ")
    /// Option Explicit
    /// ' ...
    /// ]]>
    /// </example>
    public sealed class MissingAnnotationArgumentInspection : ParseTreeInspectionBase
    {
        public MissingAnnotationArgumentInspection(RubberduckParserState state)
            : base(state) { }

        public override CodeKind TargetKindOfCode => CodeKind.AttributesCode;

        public override IInspectionListener Listener { get; } =
            new InvalidAnnotationStatementListener();

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return (from result in Listener.Contexts
                    let context = (VBAParser.AnnotationContext)result.Context 
                    where context.annotationName().GetText() == AnnotationType.Ignore.ToString() 
                       || context.annotationName().GetText() == AnnotationType.Folder.ToString() 
                    where context.annotationArgList() == null 
                    select new QualifiedContextInspectionResult(this,
                                                string.Format(InspectionResults.MissingAnnotationArgumentInspection,
                                                              ((VBAParser.AnnotationContext)result.Context).annotationName().GetText()),
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
