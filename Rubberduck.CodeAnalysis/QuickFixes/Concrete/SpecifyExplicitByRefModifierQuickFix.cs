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
    /// Introduces an explicit 'ByRef' modifier for a parameter that is implicitly passed by reference.
    /// </summary>
    /// <inspections>
    /// <inspection name="ImplicitByRefModifierInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="true" module="true" project="true" all="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething(value As Long)
    ///     Debug.Print value
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething(ByRef value As Long)
    ///     Debug.Print value
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class SpecifyExplicitByRefModifierQuickFix : QuickFixBase
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public SpecifyExplicitByRefModifierQuickFix(IDeclarationFinderProvider declarationFinderProvider)
            : base(typeof(ImplicitByRefModifierInspection))
        {
            _declarationFinderProvider = declarationFinderProvider;
        }

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            if (!(result.Target is ParameterDeclaration parameter))
            {
                return;
            }

            AddByRefIdentifier(rewriteSession, parameter);

            var finder = _declarationFinderProvider.DeclarationFinder;
            var parentDeclaration = parameter.ParentDeclaration;

            if (parentDeclaration is ModuleBodyElementDeclaration enclosingMember 
                && enclosingMember.IsInterfaceMember)
            {
                var parameterIndex = ParameterIndex(parameter, enclosingMember);
                AddByRefIdentifierToImplementations(enclosingMember, parameterIndex, finder, rewriteSession);
            }

            if (parentDeclaration is EventDeclaration enclosingEvent)
            {
                var parameterIndex = ParameterIndex(parameter, enclosingEvent);
                AddByRefIdentifierToHandlers(enclosingEvent, parameterIndex, finder, rewriteSession);
            }
        }

        private static void AddByRefIdentifierToImplementations(
            ModuleBodyElementDeclaration interfaceMember,
            int parameterIndex, 
            DeclarationFinder finder, 
            IRewriteSession rewriteSession)
        {
            var implementationParameters = finder.FindInterfaceImplementationMembers(interfaceMember)
                .Select(implementation => implementation.Parameters[parameterIndex]);

            foreach (var parameter in implementationParameters)
            {
                AddByRefIdentifier(rewriteSession, parameter);
            }
        }

        private static void AddByRefIdentifierToHandlers(
            EventDeclaration eventDeclaration,
            int parameterIndex,
            DeclarationFinder finder,
            IRewriteSession rewriteSession)
        {
            var handlers = finder.FindEventHandlers(eventDeclaration)
                .Select(implementation => implementation.Parameters[parameterIndex]);

            foreach (var parameter in handlers)
            {
                AddByRefIdentifier(rewriteSession, parameter);
            }
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.ImplicitByRefModifierQuickFix;

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;

        private static int ParameterIndex(ParameterDeclaration parameter, IParameterizedDeclaration enclosingMember)
        {
            return enclosingMember.Parameters.IndexOf(parameter);
        }

        private static void AddByRefIdentifier(IRewriteSession rewriteSession, ParameterDeclaration parameter)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(parameter.QualifiedModuleName);
            var context = (VBAParser.ArgContext) parameter.Context;
            if (context.BYREF() == null)
            {
                rewriter.InsertBefore(context.unrestrictedIdentifier().Start.TokenIndex, "ByRef ");
            }
        }
    }
}
