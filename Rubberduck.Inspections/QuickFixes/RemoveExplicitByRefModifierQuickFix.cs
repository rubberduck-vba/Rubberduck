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
                var interfaceParameterName = GetParameterIdentifierName(context);
                
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
                    var parameterName = GetParameterIdentifierName(parameterContext);

                    if (parameterName == interfaceParameterName)
                    {
                        RemoveByRefIdentifier(_state.GetRewriter(parameter), parameterContext);
                    }
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
            if (context.BYREF() != null)
            {
                rewriter.Remove(context.BYREF());
                rewriter.Remove(context.whiteSpace().First());
                // DO WYWALENIA!
                rewriter.Rewrite();
            }
        }

        private static string GetParameterIdentifierName(VBAParser.ArgContext context)
        {
            var identifier = context.unrestrictedIdentifier().identifier();
            var identifierName = identifier.untypedIdentifier() != null
                    ? identifier.untypedIdentifier().identifierValue().GetText()
                    : identifier.typedIdentifier().untypedIdentifier().identifierValue().GetText();

            return identifierName;
        }
    }
}
