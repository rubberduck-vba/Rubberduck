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
    /// Highlights implicit ByRef modifiers in user code.
    /// </summary>
    /// <why>
    /// In modern VB (VB.NET), the implicit modifier is ByVal, as it is in most other programming languages.
    /// Making the ByRef modifiers explicit can help surface potentially unexpected language defaults.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub DoSomething(foo As Long)
    ///     foo = 42
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Sub DoSomething(ByRef foo As Long)
    ///     foo = 42
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class ImplicitByRefModifierInspection : ParseTreeInspectionBase
    {
        public ImplicitByRefModifierInspection(RubberduckParserState state)
            : base(state)
        {
        }

        public override IInspectionListener Listener { get; } = new ImplicitByRefModifierListener();

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var builtInEventHandlerContexts = State.DeclarationFinder.FindEventHandlers().Select(handler => handler.Context).ToHashSet();
            var interfaceImplementationMemberContexts = State.DeclarationFinder.FindAllInterfaceImplementingMembers().Select(member => member.Context).ToHashSet();

            var issues = Listener.Contexts.Where(context =>
                !builtInEventHandlerContexts.Contains(context.Context.Parent.Parent) &&
                !interfaceImplementationMemberContexts.Contains(context.Context.Parent.Parent));

            return issues.Select(issue =>
            {
                var identifier = ((VBAParser.ArgContext)issue.Context)
                    .unrestrictedIdentifier()
                    .identifier();

                return new QualifiedContextInspectionResult(this,
                    string.Format(InspectionResults.ImplicitByRefModifierInspection,
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
