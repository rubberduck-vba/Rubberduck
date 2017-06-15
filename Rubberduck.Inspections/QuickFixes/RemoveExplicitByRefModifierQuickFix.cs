using System;
using Rubberduck.Parsing.Grammar;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.PostProcessing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public class RemoveExplicitByRefModifierQuickFix : IQuickFix
    {
        private readonly RubberduckParserState _state;

        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type>
        {
            typeof(RedundantByRefModifierInspection)
        };

        public RemoveExplicitByRefModifierQuickFix(RubberduckParserState state)
        {
            _state = state;
        }

        public IReadOnlyCollection<Type> SupportedInspections => _supportedInspections.ToList();

        public void Fix(IInspectionResult result)
        {
            var context = (VBAParser.ArgContext) result.Context;

            RemoveByRefIdentifier(_state.GetRewriter(result.QualifiedSelection.QualifiedName), context);

            var interfaceMembers = _state.DeclarationFinder.FindAllInterfaceMembers().ToArray();

            var matchingInterfaceMemberContext = interfaceMembers.Select(member => member.Context).FirstOrDefault(c => c == context.Parent.Parent);

            if (matchingInterfaceMemberContext != null)
            {
                var implementationMembers =
                    _state.AllUserDeclarations.FindInterfaceImplementationMembers(interfaceMembers.First(
                        member => member.Context == matchingInterfaceMemberContext)).ToArray();

                var parameters =
                    _state.DeclarationFinder.UserDeclarations(DeclarationType.Parameter)
                        .Where(p => implementationMembers.Contains(p.ParentDeclaration))
                        .Cast<ParameterDeclaration>()
                        .ToArray();

                foreach (var parameter in parameters)
                {
                    RemoveByRefIdentifier(_state.GetRewriter(parameter), (VBAParser.ArgContext) parameter.Context);
                }
            }
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.RedundantByRefModifierQuickFix;
        }

        public bool CanFixInProcedure => true;
        public bool CanFixInModule => true;
        public bool CanFixInProject => true;

        private static void RemoveByRefIdentifier(IModuleRewriter rewriter, VBAParser.ArgContext context)
        {
            rewriter.Remove(context.BYREF());
            rewriter.Remove(context.whiteSpace().First());
            rewriter.Rewrite();
        }
    }
}
