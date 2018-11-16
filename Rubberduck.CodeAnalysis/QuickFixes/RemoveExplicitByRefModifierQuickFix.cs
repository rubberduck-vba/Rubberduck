using System;
using Rubberduck.Parsing.Grammar;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Extensions;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class RemoveExplicitByRefModifierQuickFix : QuickFixBase
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public RemoveExplicitByRefModifierQuickFix(IDeclarationFinderProvider declarationFinderProvider)
            : base(typeof(RedundantByRefModifierInspection))
        {
            _declarationFinderProvider = declarationFinderProvider;
        }

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var context = (VBAParser.ArgContext) result.Context;

            RemoveByRefIdentifier(rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName), context);

            var interfaceMembers = _declarationFinderProvider.DeclarationFinder.FindAllInterfaceMembers().ToArray();

            var matchingInterfaceMemberContext = interfaceMembers.Select(member => member.Context).FirstOrDefault(c => c == context.Parent.Parent);

            if (matchingInterfaceMemberContext != null)
            {
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
                    var parameterContext = (VBAParser.ArgContext) parameter.Context;
                    var parameterIndex = GetParameterIndex(parameterContext);

                    if (parameterIndex == interfaceParameterIndex)
                    {
                        RemoveByRefIdentifier(rewriteSession.CheckOutModuleRewriter(parameter.QualifiedModuleName), parameterContext);
                    }
                }
            }
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.RedundantByRefModifierQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;

        private static int GetParameterIndex(VBAParser.ArgContext context)
        {
            return Array.IndexOf(((VBAParser.ArgListContext)context.Parent).arg().ToArray(), context);
        }

        private static void RemoveByRefIdentifier(IModuleRewriter rewriter, VBAParser.ArgContext context)
        {
            if (context.BYREF() != null)
            {
                rewriter.Remove(context.BYREF());
                rewriter.Remove(context.whiteSpace().First());
            }
        }
    }
}
