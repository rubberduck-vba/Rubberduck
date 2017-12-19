using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Concrete
{
    [CannotAnnotate]
    public sealed class MissingAttributeInspection : ParseTreeInspectionBase
    {
        public MissingAttributeInspection(RubberduckParserState state)
            : base(state)
        {
            Listener = new MissingMemberAttributeListener(state);
        }

        public override ParsePass Pass => ParsePass.AttributesPass;

        public override CodeInspectionType InspectionType => CodeInspectionType.RubberduckOpportunities;
        public override IInspectionListener Listener { get; }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            return Listener.Contexts.Select(context =>
            {
                var name = string.Format(InspectionsUI.MissingAttributeInspectionResultFormat, context.MemberName.MemberName,
                    ((VBAParser.AnnotationContext) context.Context).annotationName().GetText());
                return new QualifiedContextInspectionResult(this, name, context);
            });
        }

        public class MissingMemberAttributeListener : ParseTreeListeners.AttributeAnnotationListener
        {
            public MissingMemberAttributeListener(RubberduckParserState state) : base(state) { }

            public override void ExitAnnotation(VBAParser.AnnotationContext context)
            {
                var annotationType = context.AnnotationType;

                if (!annotationType.HasFlag(AnnotationType.Attribute))
                {
                    return;
                }

                var isMemberAnnotation = annotationType.HasFlag(AnnotationType.MemberAnnotation);
                var isModuleScope = CurrentScopeDeclaration.DeclarationType.HasFlag(DeclarationType.Module);

                if (isModuleScope && !isMemberAnnotation)
                {
                    // module-level annotation
                    var module = State.DeclarationFinder.UserDeclarations(DeclarationType.Module).Single(m => m.QualifiedName.QualifiedModuleName.Equals(CurrentModuleName));
                    if (!module.Attributes.HasAttributeFor(context.AnnotationType))
                    {
                        AddContext(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
                    }
                }
                else if (isMemberAnnotation)
                {
                    // member-level annotation is above the context for the first member in the module..
                    if (isModuleScope)
                    {
                        CurrentScopeDeclaration = FirstMember;
                    }

                    var member = Members.Value.Single(m => m.Key.Equals(CurrentScopeDeclaration.QualifiedName.MemberName));
                    if (!member.Value.Attributes.HasAttributeFor(context.AnnotationType, member.Key))
                    {
                        AddContext(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
                    }
                }
                else
                {
                    // annotation is illegal. ignore.
                }
            }
        }
    }
}