using System.Diagnostics;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Identifies parameter declarations that are not used.
    /// </summary>
    /// <why>
    /// Declarations that are not used anywhere should probably be removed.
    /// </why>
    /// <remarks>
    /// Not all unused parameters can/should be removed: ignore any inspection results for 
    /// event handler procedures and interface members that Rubberduck isn't recognizing as such.
    /// </remarks>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething(ByVal foo As Long, ByVal bar As Long)
    ///     Debug.Print foo
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Option Explicit
    /// Public Sub DoSomething(ByVal foo As Long, ByVal bar As Long)
    ///     Debug.Print foo, bar
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class ParameterNotUsedInspection : DeclarationInspectionBase
    {
        public ParameterNotUsedInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider, DeclarationType.Parameter)
        {}

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            if (declaration.References.Any()
                || !(declaration is ParameterDeclaration parameter))
            {
                return false;
            }

            var enclosingMember = parameter.ParentDeclaration;
            if (IsLibraryMethod(enclosingMember))
            {
                return false;
            }

            if (enclosingMember is EventDeclaration eventDeclaration)
            {
                return ThereAreHandlersAndNoneUsesTheParameter(parameter, eventDeclaration, finder);
            }

            if (enclosingMember is ModuleBodyElementDeclaration member)
            {
                if (member.IsInterfaceMember)
                {
                    return ThereAreImplementationsAndNoneUsesTheParameter(parameter, member, finder);
                }

                if (member.IsInterfaceImplementation
                    || finder.FindEventHandlers().Contains(member))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool IsLibraryMethod(Declaration declaration)
        {
            return declaration.DeclarationType == DeclarationType.LibraryProcedure
                   || declaration.DeclarationType == DeclarationType.LibraryFunction;
        }

        private static bool ThereAreImplementationsAndNoneUsesTheParameter(ParameterDeclaration parameter, ModuleBodyElementDeclaration interfaceMember, DeclarationFinder finder)
        {
            if (!TryFindParameterIndex(parameter, interfaceMember, out var parameterIndex))
            {
                //This really should never happen.
                Debug.Fail($"Could not find index for parameter {parameter.IdentifierName} in interface member {interfaceMember.IdentifierName}.");
                return false;
            }

            var implementations = finder.FindInterfaceImplementationMembers(interfaceMember).ToList();

            //We do not want to report all parameters of not implemented interfaces.
            return implementations.Any() 
                   && implementations.All(implementation => ParameterAtIndexIsNotUsed(implementation, parameterIndex));
        }

        private static bool TryFindParameterIndex(ParameterDeclaration parameter, IParameterizedDeclaration enclosingMember, out int parameterIndex)
        {
            parameterIndex = enclosingMember.Parameters
                .ToList()
                .IndexOf(parameter);
            return parameterIndex != -1;
        }

        private static bool ParameterAtIndexIsNotUsed(IParameterizedDeclaration declaration, int parameterIndex)
        {
            var parameter = declaration?.Parameters.ElementAtOrDefault(parameterIndex);
            return parameter != null
                   && !parameter.References.Any();
        }

        private static bool ThereAreHandlersAndNoneUsesTheParameter(ParameterDeclaration parameter, EventDeclaration eventDeclaration, DeclarationFinder finder)
        {
            if (!TryFindParameterIndex(parameter, eventDeclaration, out var parameterIndex))
            {
                //This really should never happen.
                Debug.Fail($"Could not find index for parameter {parameter.IdentifierName} in event {eventDeclaration.IdentifierName}.");
                return false;
            }

            if (!eventDeclaration.IsUserDefined)
            {
                return false;
            }

            var handlers = finder.FindEventHandlers(eventDeclaration).ToList();

            //We do not want to report all parameters of not handled events.
            return handlers.Any() 
                   && handlers.All(handler => ParameterAtIndexIsNotUsed(handler, parameterIndex));
        }

        protected override string ResultDescription(Declaration declaration)
        {
            return string.Format(InspectionResults.ParameterNotUsedInspection, declaration.IdentifierName).Capitalize();
        }
    }
}
