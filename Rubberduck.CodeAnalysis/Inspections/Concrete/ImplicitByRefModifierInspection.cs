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
        {}

        public override IInspectionListener Listener { get; } = new ImplicitByRefModifierListener();
        protected override string ResultDescription(QualifiedContext<ParserRuleContext> context)
        {
            var identifier = ((VBAParser.ArgContext)context.Context)
                .unrestrictedIdentifier()
                .identifier();

            var identifierText = identifier.untypedIdentifier() != null
                ? identifier.untypedIdentifier().identifierValue().GetText()
                : identifier.typedIdentifier().untypedIdentifier().identifierValue().GetText();

            return string.Format(
                InspectionResults.ImplicitByRefModifierInspection,
                identifierText);
        }

        protected override bool IsResultContext(QualifiedContext<ParserRuleContext> context)
        {
            //FIXME : This should really be a declaration inspection on the parameter. 
            var builtInEventHandlerContexts = State.DeclarationFinder.FindEventHandlers().Select(handler => handler.Context).ToHashSet();
            var interfaceImplementationMemberContexts = State.DeclarationFinder.FindAllInterfaceImplementingMembers().Select(member => member.Context).ToHashSet();

            return !builtInEventHandlerContexts.Contains(context.Context.Parent.Parent)
                   && !interfaceImplementationMemberContexts.Contains(context.Context.Parent.Parent);
        }

        public class ImplicitByRefModifierListener : InspectionListenerBase
        {
            public override void ExitArg(VBAParser.ArgContext context)
            {
                if (context.PARAMARRAY() == null && context.BYVAL() == null && context.BYREF() == null)
                {
                    SaveContext(context);
                }
            }
        }
    }
}
