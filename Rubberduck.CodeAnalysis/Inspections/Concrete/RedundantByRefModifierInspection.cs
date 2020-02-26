using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
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
        protected override string ResultDescription(QualifiedContext<ParserRuleContext> context)
        {
            var identifier = ((VBAParser.ArgContext)context.Context)
                .unrestrictedIdentifier()
                .identifier();

            var identifierText = identifier.untypedIdentifier() != null
                ? identifier.untypedIdentifier().identifierValue().GetText()
                : identifier.typedIdentifier().untypedIdentifier().identifierValue().GetText();

            return string.Format(
                InspectionResults.RedundantByRefModifierInspection,
                identifierText);
        }

        protected override bool IsResultContext(QualifiedContext<ParserRuleContext> context)
        {
            //FIXME This should be an inspection on parameter declarations.
            var finder = DeclarationFinderProvider.DeclarationFinder;
            var builtInEventHandlerContexts = finder.FindEventHandlers().Select(handler => handler.Context).ToHashSet();
            var interfaceImplementationMemberContexts = finder.FindAllInterfaceImplementingMembers().Select(member => member.Context).ToHashSet();

            return !builtInEventHandlerContexts.Contains(context.Context.Parent.Parent)
                   && !interfaceImplementationMemberContexts.Contains(context.Context.Parent.Parent);
        }

        public class RedundantByRefModifierListener : InspectionListenerBase
        {
            public override void ExitArg(VBAParser.ArgContext context)
            {
                if (context.BYREF() != null)
                {
                   SaveContext(context);
                }
            }
        }
    }
}
