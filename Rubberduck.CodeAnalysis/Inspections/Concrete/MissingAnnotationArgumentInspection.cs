using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Warns about a malformed Rubberduck annotation that is missing an argument.
    /// </summary>
    /// <why>
    /// Some annotations require arguments; if the argument isn't specified, the annotation is nothing more than an obscure comment.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// '@Folder
    /// '@ModuleDescription
    /// Option Explicit
    /// ' ...
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// '@Folder("MyProject.XYZ")
    /// '@ModuleDescription("This module does XYZ")
    /// Option Explicit
    /// ' ...
    /// ]]>
    /// </example>
    public sealed class MissingAnnotationArgumentInspection : ParseTreeInspectionBase
    {
        private readonly RubberduckParserState _state;

        public MissingAnnotationArgumentInspection(RubberduckParserState state)
            : base(state)
        {
            _state = state;
        }

        public override CodeKind TargetKindOfCode => CodeKind.AttributesCode;

        public override IInspectionListener Listener { get; } =
            new InvalidAnnotationStatementListener();

        protected override string ResultDescription(QualifiedContext<ParserRuleContext> context)
        {
            var expressionText = ((VBAParser.AnnotationContext) context.Context).annotationName().GetText();
            return string.Format(
                InspectionResults.MissingAnnotationArgumentInspection,
                expressionText);
        }

        protected override bool IsResultContext(QualifiedContext<ParserRuleContext> context)
        {
            // FIXME don't actually use listeners here, iterate the Annotations instead
            // FIXME don't maintain a separate list for annotations that require arguments, instead use AnnotationAttribute to store that information
            var annotationContext = (VBAParser.AnnotationContext) context.Context;
            return (annotationContext.annotationName().GetText() == "Ignore"
                    || annotationContext.annotationName().GetText() == "Folder")
                    && annotationContext.annotationArgList() == null;
        }

        public class InvalidAnnotationStatementListener : InspectionListenerBase
        {
            public override void ExitAnnotation(VBAParser.AnnotationContext context)
            {
                if (context.annotationName() != null)
                {
                    SaveContext(context);
                }
            }
        }
    }
}
