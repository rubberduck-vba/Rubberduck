using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.CodeAnalysis.Inspections.Extensions;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Identifies obsolete 16-bit integer variables.
    /// </summary>
    /// <why>
    /// Modern processors are optimized for processing 32-bit integers; internally, a 16-bit integer is still stored as a 32-bit value.
    /// Unless code is interacting with APIs that require a 16-bit integer, a Long (32-bit integer) should be used instead.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim rowCount As Integer
    ///     rowCount = Sheet1.Rows.Count ' overflow: maximum 16-bit signed integer value is only 32,767 (2^15-1)!
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim rowCount As Long
    ///     rowCount = Sheet1.Rows.Count ' all good: maximum 32-bit signed integer value is 2,147,483,647 (2^31-1)!
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class IntegerDataTypeInspection : DeclarationInspectionBase
    {
        public IntegerDataTypeInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {}

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            if (declaration.AsTypeName != Tokens.Integer)
            {
                return false;
            }

            switch (declaration)
            {
                case ParameterDeclaration parameter:
                    return ParameterIsResult(parameter, finder);
                case ModuleBodyElementDeclaration member:
                    return MethodIsResult(member);
                default:
                    return declaration.DeclarationType != DeclarationType.LibraryFunction;
            }
        }

        private static bool ParameterIsResult(ParameterDeclaration parameter, DeclarationFinder finder)
        {
            var enclosingMember = parameter.ParentDeclaration;
            if (!(enclosingMember is ModuleBodyElementDeclaration member))
            {
                return false;
            }

            return !member.IsInterfaceImplementation
                   && member.DeclarationType != DeclarationType.LibraryFunction
                   && member.DeclarationType != DeclarationType.LibraryProcedure
                   && !finder.FindEventHandlers().Contains(member);
        }

        private static bool MethodIsResult(ModuleBodyElementDeclaration member)
        {
            return !member.IsInterfaceImplementation;
        }

        protected override string ResultDescription(Declaration declaration)
        {
            var declarationType = declaration.DeclarationType.ToLocalizedString();
            var declarationName = declaration.IdentifierName;
            return string.Format(
                Resources.Inspections.InspectionResults.IntegerDataTypeInspection,
                declarationType,
                declarationName);
        }
    }
}
