using System.Linq;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.InternalApi.Extensions;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Removes an explicit ByRef modifier, making it implicit.
    /// </summary>
    /// <inspections>
    /// <inspection name="RedundantByRefModifierInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="true" module="true" project="true" all="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething(ByRef value As Long)
    ///     '...
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething(value As Long)
    ///     '...
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class RemoveExplicitByRefModifierQuickFix : QuickFixBase
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public RemoveExplicitByRefModifierQuickFix(IDeclarationFinderProvider declarationFinderProvider)
            : base(typeof(RedundantByRefModifierInspection))
        {
            _declarationFinderProvider = declarationFinderProvider;
        }

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            if (!(result.Target is ParameterDeclaration parameter))
            {
                return;
            }

            RemoveByRefIdentifier(rewriteSession, parameter);

            var finder = _declarationFinderProvider.DeclarationFinder;
            var parentDeclaration = parameter.ParentDeclaration;

            if (parentDeclaration is ModuleBodyElementDeclaration enclosingMember
                && enclosingMember.IsInterfaceMember)
            {
                var parameterIndex = ParameterIndex(parameter, enclosingMember);
                RemoveByRefIdentifierFromImplementations(enclosingMember, parameterIndex, finder, rewriteSession);
            }

            if (parentDeclaration is EventDeclaration enclosingEvent)
            {
                var parameterIndex = ParameterIndex(parameter, enclosingEvent);
                RemoveByRefIdentifierFromHandlers(enclosingEvent, parameterIndex, finder, rewriteSession);
            }
        }

        private static void RemoveByRefIdentifierFromImplementations(
            ModuleBodyElementDeclaration interfaceMember,
            int parameterIndex,
            DeclarationFinder finder,
            IRewriteSession rewriteSession)
        {
            var implementationParameters = finder.FindInterfaceImplementationMembers(interfaceMember)
                .Select(implementation => implementation.Parameters[parameterIndex]);

            foreach (var parameter in implementationParameters)
            {
                RemoveByRefIdentifier(rewriteSession, parameter);
            }
        }

        private static void RemoveByRefIdentifierFromHandlers(
            EventDeclaration eventDeclaration,
            int parameterIndex,
            DeclarationFinder finder,
            IRewriteSession rewriteSession)
        {
            var handlers = finder.FindEventHandlers(eventDeclaration)
                .Select(implementation => implementation.Parameters[parameterIndex]);

            foreach (var parameter in handlers)
            {
                RemoveByRefIdentifier(rewriteSession, parameter);
            }
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.RedundantByRefModifierQuickFix;

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;

        private static int ParameterIndex(ParameterDeclaration parameter, IParameterizedDeclaration enclosingMember)
        {
            return enclosingMember.Parameters.IndexOf(parameter);
        }

        private static void RemoveByRefIdentifier(IRewriteSession rewriteSession, ParameterDeclaration parameter)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(parameter.QualifiedModuleName);
            var context = (VBAParser.ArgContext)parameter.Context;

            if (context.BYREF() != null)
            {
                rewriter.Remove(context.BYREF());
                rewriter.Remove(context.whiteSpace().First());
            }
        }
    }
}
