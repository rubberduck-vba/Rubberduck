using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Inspections
{
    public sealed class EmptyStringLiteralInspection : InspectionBase, IParseTreeInspection
    {
        public EmptyStringLiteralInspection(RubberduckParserState state)
            : base(state)
        {
        }

        public override string Meta { get { return InspectionsUI.EmptyStringLiteralInspectionMeta; } }
        public override string Description { get { return InspectionsUI.EmptyStringLiteralInspection; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }

        public ParseTreeResults ParseTreeResults { get; set; }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {   
            if (ParseTreeResults == null)
            {
                return new InspectionResultBase[] { };
            }
            return ParseTreeResults.EmptyStringLiterals
                .Where(s => !HasIgnoreEmptyStringLiteralAnnotations(s.ModuleName.Component, s.Context.Start.Line))
                .Select(context => new EmptyStringLiteralInspectionResult(this,
                            new QualifiedContext<ParserRuleContext>(context.ModuleName, context.Context)));
        }

        private bool HasIgnoreEmptyStringLiteralAnnotations(VBComponent component, int line)
        {
            var annotations = State.GetModuleAnnotations(component).ToList();

            if (State.GetModuleAnnotations(component) == null)
            {
                return false;
            }
            
            // VBE 1-based indexing
            for (var i = line - 1; i >= 1; i--)
            {
                var annotation = annotations.SingleOrDefault(a => a.QualifiedSelection.Selection.StartLine == i) as IgnoreAnnotation;
                if (annotation != null && annotation.InspectionNames.Contains(AnnotationName))
                {
                    return true;
                }
            }

            return false;
        }

        public class EmptyStringLiteralListener : VBAParserBaseListener
        {
            private readonly IList<VBAParser.LiteralExpressionContext> _contexts = new List<VBAParser.LiteralExpressionContext>();
            public IEnumerable<VBAParser.LiteralExpressionContext> Contexts { get { return _contexts; } }

            public override void ExitLiteralExpression(VBAParser.LiteralExpressionContext context)
            {
                var literal = context.STRINGLITERAL();
                if (literal != null && literal.GetText() == "\"\"")
                {
                    _contexts.Add(context);
                }
            }
        }
    }
}
