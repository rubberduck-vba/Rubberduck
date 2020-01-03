using System;
using Rubberduck.Parsing.Grammar;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.JunkDrawer.Extensions;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    /// <summary>
    /// Introduces an explicit 'ByRef' modifier for a parameter that is implicitly passed by reference.
    /// </summary>
    /// <inspections>
    /// <inspection name="ImplicitByRefModifierInspection" />
    /// </inspections>
    /// <canfix procedure="true" module="true" project="true" />
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
    public sealed class SpecifyExplicitByRefModifierQuickFix : QuickFixBase
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public SpecifyExplicitByRefModifierQuickFix(IDeclarationFinderProvider declarationFinderProvider)
            : base(typeof(ImplicitByRefModifierInspection))
        {
            _declarationFinderProvider = declarationFinderProvider;
        }

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var context = (VBAParser.ArgContext)result.Context;

            AddByRefIdentifier(rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName), context);

            var interfaceMembers = _declarationFinderProvider.DeclarationFinder.FindAllInterfaceMembers().ToArray();

            var matchingInterfaceMemberContext = interfaceMembers.Select(member => member.Context).FirstOrDefault(c => c == context.Parent.Parent);

            if (matchingInterfaceMemberContext == null)
            {
                return;
            }
            
            var interfaceParameterIndex = GetParameterIndex(context);

            var implementationMembers =
                _declarationFinderProvider.DeclarationFinder.FindInterfaceImplementationMembers(interfaceMembers.First(
                    member => member.Context == matchingInterfaceMemberContext)).ToHashSet();

            var parameters =
                _declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.Parameter)
                    .Where(p => implementationMembers.Contains(p.ParentDeclaration))
                    .Cast<ParameterDeclaration>()
                    .ToArray();

            foreach (var parameter in parameters)
            {
                var parameterContext = (VBAParser.ArgContext)parameter.Context;
                var parameterIndex = GetParameterIndex(parameterContext);

                if (parameterIndex == interfaceParameterIndex)
                {
                    AddByRefIdentifier(rewriteSession.CheckOutModuleRewriter(parameter.QualifiedModuleName), parameterContext);
                }
            }
            
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.ImplicitByRefModifierQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;

        private static int GetParameterIndex(VBAParser.ArgContext context)
        {
            return Array.IndexOf(((VBAParser.ArgListContext)context.Parent).arg().ToArray(), context);
        }

        private static void AddByRefIdentifier(IModuleRewriter rewriter, VBAParser.ArgContext context)
        {
            if (context.BYREF() == null)
            {
                rewriter.InsertBefore(context.unrestrictedIdentifier().Start.TokenIndex, "ByRef ");
            }
        }
    }
}
