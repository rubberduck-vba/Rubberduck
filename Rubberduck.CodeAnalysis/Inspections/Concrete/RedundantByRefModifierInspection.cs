using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Extensions;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class RedundantByRefModifierInspection : ParseTreeInspectionBase
    {
        public RedundantByRefModifierInspection(RubberduckParserState state)
            : base(state)
        {
        }

        public override IInspectionListener Listener { get; } = new RedundantByRefModifierListener();

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var builtInEventHandlerContexts = State.DeclarationFinder.FindEventHandlers().Select(handler => handler.Context).ToHashSet();
            var interfaceImplementationMemberContexts = State.DeclarationFinder.FindAllInterfaceImplementingMembers().Select(member => member.Context).ToHashSet();

            var issues = Listener.Contexts.Where(context =>
                !IsIgnoringInspectionResultFor(context.ModuleName, context.Context.Start.Line) &&
                !builtInEventHandlerContexts.Contains(context.Context.Parent.Parent) &&
                !interfaceImplementationMemberContexts.Contains(context.Context.Parent.Parent));

            return issues.Select(issue =>
            {
                var identifier = ((VBAParser.ArgContext) issue.Context)
                    .unrestrictedIdentifier()
                    .identifier();

                return new QualifiedContextInspectionResult(this,
                    string.Format(InspectionResults.RedundantByRefModifierInspection,
                        identifier.untypedIdentifier() != null
                            ? identifier.untypedIdentifier().identifierValue().GetText()
                            : identifier.typedIdentifier().untypedIdentifier().identifierValue().GetText()), issue);
            });
        }

        public class RedundantByRefModifierListener : VBAParserBaseListener, IInspectionListener
        {
            private readonly List<QualifiedContext<ParserRuleContext>> _contexts = new List<QualifiedContext<ParserRuleContext>>();

            public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts  => _contexts;

            public QualifiedModuleName CurrentModuleName { get; set; }

            public void ClearContexts()
            {
                _contexts.Clear();
            }

            public override void ExitArg(VBAParser.ArgContext context)
            {
                if (context.BYREF() != null)
                {
                    _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
                }
            }
        }
    }
}
