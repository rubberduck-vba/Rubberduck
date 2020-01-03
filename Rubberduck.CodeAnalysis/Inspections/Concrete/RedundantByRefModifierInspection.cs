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
using Rubberduck.Inspections.Inspections.Extensions;
using Rubberduck.JunkDrawer.Extensions;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Identifies redundant ByRef modifiers.
    /// </summary>
    /// <why>
    /// Out of convention or preference, explicit ByRef modifiers could be considered redundant since they are the implicit default. 
    /// This inspection can ensure the consistency of the convention.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething(ByRef foo As Long)
    ///     foo = foo + 17
    ///     Debug.Print foo
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Option Explicit
    /// Public Sub DoSomething(foo As Long)
    ///     foo = foo + 17
    ///     Debug.Print foo
    /// End Sub
    /// ]]>
    /// </example>
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
