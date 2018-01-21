using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class ImplicitByRefModifierInspection : ParseTreeInspectionBase
    {
        public ImplicitByRefModifierInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Hint)
        {
        }

        public override CodeInspectionType InspectionType => CodeInspectionType.CodeQualityIssues;

        public override IInspectionListener Listener { get; } = new ImplicitByRefModifierListener();

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var builtInEventHandlerContexts = State.DeclarationFinder.FindEventHandlers().Select(handler => handler.Context).ToHashSet();
            var interfaceImplementationMemberContexts = UserDeclarations.FindInterfaceImplementationMembers().Select(member => member.Context).ToHashSet();

            var issues = Listener.Contexts.Where(context =>
                !IsIgnoringInspectionResultFor(context.ModuleName, context.Context.Start.Line) &&
                !builtInEventHandlerContexts.Contains(context.Context.Parent.Parent) &&
                !interfaceImplementationMemberContexts.Contains(context.Context.Parent.Parent));

            return issues.Select(issue =>
            {
                var identifier = ((VBAParser.ArgContext)issue.Context)
                    .unrestrictedIdentifier()
                    .identifier();

                return new QualifiedContextInspectionResult(this,
                    string.Format(InspectionsUI.ImplicitByRefModifierInspectionResultFormat,
                        identifier.untypedIdentifier() != null
                            ? identifier.untypedIdentifier().identifierValue().GetText()
                            : identifier.typedIdentifier().untypedIdentifier().identifierValue().GetText()), issue);
            });
        }

        public class ImplicitByRefModifierListener : VBAParserBaseListener, IInspectionListener
        {
            private readonly List<QualifiedContext<ParserRuleContext>> _contexts = new List<QualifiedContext<ParserRuleContext>>();

            public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;

            public QualifiedModuleName CurrentModuleName { get; set; }

            public void ClearContexts()
            {
                _contexts.Clear();
            }

            public override void ExitArg(VBAParser.ArgContext context)
            {
                if (context.PARAMARRAY() == null && context.BYVAL() == null && context.BYREF() == null)
                {
                    _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
                }
            }
        }
    }
}
