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

namespace Rubberduck.Inspections.Concrete
{
    public sealed class MissingAnnotationInspection : ParseTreeInspectionBase
    {
        public MissingAnnotationInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion)
        {
            Listener = new MissingAnnotationListener(state);
        }

        public override CodeInspectionType InspectionType => CodeInspectionType.RubberduckOpportunities;
        public override IInspectionListener Listener { get; }
        public override ParsePass Pass => ParsePass.AttributesPass;

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return Listener.Contexts.Select(context =>
            {
                var member = string.IsNullOrEmpty(context.MemberName.MemberName)
                    ? context.ModuleName.Name
                    : context.MemberName.MemberName;

                var name = string.Format(InspectionsUI.MissingAnnotationInspectionResultFormat, 
                    member, ((VBAParser.AttributeStmtContext) context.Context).AnnotationType().ToString());

                return new QualifiedContextInspectionResult(this, name, context);
            });
        }

        public class MissingAnnotationListener : ParseTreeListeners.AttributeAnnotationListener
        {
            public MissingAnnotationListener(RubberduckParserState state) : base(state) { }
            
            public override void ExitAttributeStmt(VBAParser.AttributeStmtContext context)
            {
                var annotations = CurrentScopeDeclaration?.Annotations;

                var type = context.AnnotationType();
                if (type != null && (annotations?.All(a => a.AnnotationType != type) ?? false))
                {
                    if (type.Value.HasFlag(AnnotationType.ModuleAnnotation))
                    {
                        // attribute is mapped to an annotation, but current scope doesn't have that annotation:
                        AddContext(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
                    }

                    if(type.Value.HasFlag(AnnotationType.MemberAnnotation))
                    {
                        AddContext(new QualifiedContext<ParserRuleContext>(CurrentScopeDeclaration.QualifiedName, context));
                    }
                }
            }
        }
    }
}