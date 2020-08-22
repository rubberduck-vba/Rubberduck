using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Warns about implicit local variables that are used but never declared.
    /// </summary>
    /// <why>
    /// If this code compiles, then Option Explicit is omitted and compile-time validation is easily forfeited, even accidentally (e.g. typos). 
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     foo = 42 ' foo is not declared
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim foo As Long
    ///     foo = 42
    ///     ' ...
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class UndeclaredVariableInspection : DeclarationInspectionBase
    {
        public UndeclaredVariableInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider, DeclarationType.Variable)
        {}

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            return declaration.IsUndeclared;
        }

        protected override string ResultDescription(Declaration declaration)
        {
            return string.Format(InspectionResults.UndeclaredVariableInspection, declaration.IdentifierName);
        }
    }
}