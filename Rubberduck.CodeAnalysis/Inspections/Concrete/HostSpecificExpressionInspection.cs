using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Warns about host-evaluated square-bracketed expressions.
    /// </summary>
    /// <why>
    /// Host-evaluated expressions should be implementable using the host application's object model.
    /// If the expression yields an object, member calls against that object are late-bound.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     [A1].Value = 42
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     ActiveSheet.Range("A1").Value = 42
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class HostSpecificExpressionInspection : DeclarationInspectionBase
    {
        public HostSpecificExpressionInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider, DeclarationType.BracketedExpression)
        {}

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            return true;
        }

        protected override string ResultDescription(Declaration declaration)
        {
            return string.Format(InspectionResults.HostSpecificExpressionInspection, declaration.IdentifierName);
        }
    }
}