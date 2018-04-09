using System;
using Rubberduck.Parsing.Grammar;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public class RemoveExplicitByRefModifierQuickFix : QuickFixBase
    {
        private readonly RubberduckParserState _state;

        public RemoveExplicitByRefModifierQuickFix(RubberduckParserState state)
            : base(typeof(RedundantByRefModifierInspection))
        {
            _state = state;
        }

        public override void Fix(IInspectionResult result)
        {
            var context = (VBAParser.ArgContext) result.Context;

            RemoveByRefIdentifier(_state.GetRewriter(result.QualifiedSelection.QualifiedName), context);

            var interfaceMembers = _state.DeclarationFinder.FindAllInterfaceMembers().ToArray();

            var matchingInterfaceMemberContext = interfaceMembers.Select(member => member.Context).FirstOrDefault(c => c == context.Parent.Parent);

            if (matchingInterfaceMemberContext != null)
            {
                var interfaceParameterIndex = GetParameterIndex(context);

                var implementationMembers =
                    _state.AllUserDeclarations.FindInterfaceImplementationMembers(interfaceMembers.First(
                        member => member.Context == matchingInterfaceMemberContext)).ToHashSet();

                var parameters =
                    _state.DeclarationFinder.UserDeclarations(DeclarationType.Parameter)
                        .Where(p => implementationMembers.Contains(p.ParentDeclaration))
                        .Cast<ParameterDeclaration>()
                        .ToArray();

                foreach (var parameter in parameters)
                {
                    var parameterContext = (VBAParser.ArgContext) parameter.Context;
                    var parameterIndex = GetParameterIndex(parameterContext);

                    if (parameterIndex == interfaceParameterIndex)
                    {
                        RemoveByRefIdentifier(_state.GetRewriter(parameter), parameterContext);
                    }
                }
            }
        }

        public override string Description(IInspectionResult result) => InspectionsUI.RedundantByRefModifierQuickFix;

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
