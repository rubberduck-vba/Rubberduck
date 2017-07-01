using System;
using Rubberduck.Parsing.Grammar;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public class SpecifyExplicitByRefModifierQuickFix : IQuickFix
    {
        private readonly RubberduckParserState _state;

        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type>
        {
            typeof(ImplicitByRefModifierInspection)
        };

        public SpecifyExplicitByRefModifierQuickFix(RubberduckParserState state)
        {
            _state = state;
        }

        public IReadOnlyCollection<Type> SupportedInspections => _supportedInspections.ToList();

        public void Fix(IInspectionResult result)
        {
            var context = (VBAParser.ArgContext)result.Context;

            AddByRefIdentifier(_state.GetRewriter(result.QualifiedSelection.QualifiedName), context);

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
                    var parameterContext = (VBAParser.ArgContext)parameter.Context;
                    var parameterIndex = GetParameterIndex(parameterContext);

                    if (parameterIndex == interfaceParameterIndex)
                    {
                        AddByRefIdentifier(_state.GetRewriter(parameter), parameterContext);
                    }
                }
            }
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.ImplicitByRefModifierQuickFix;
        }

        public bool CanFixInProcedure => true;
        public bool CanFixInModule => true;
        public bool CanFixInProject => true;

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
