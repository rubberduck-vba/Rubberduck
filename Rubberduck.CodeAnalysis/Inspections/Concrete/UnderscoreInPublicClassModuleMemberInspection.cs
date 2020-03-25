using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Warns about public class members with an underscore in their names.
    /// </summary>
    /// <why>
    /// The public interface of any class module can be implemented by any other class module; if the public interface 
    /// contains names with underscores, other classes cannot implement it - the code will not compile. Avoid underscores; prefer PascalCase names.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Class Module">
    /// <![CDATA[
    /// '@Interface
    /// 
    /// Public Sub Do_Something() ' underscore in name makes the interface non-implementable.
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Class Module">
    /// <![CDATA[
    /// '@Interface
    /// 
    /// Public Sub DoSomething() ' PascalCase identifiers are never a problem.
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class UnderscoreInPublicClassModuleMemberInspection : DeclarationInspectionBase
    {
        public UnderscoreInPublicClassModuleMemberInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider, DeclarationType.Member)
        {}

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            return declaration.IdentifierName.Contains("_") 
                   && (declaration.Accessibility == Accessibility.Public 
                       || declaration.Accessibility == Accessibility.Implicit) 
                   && declaration.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.ClassModule)
                   && !finder.FindEventHandlers().Contains(declaration)
                   && !(declaration is ModuleBodyElementDeclaration member && member.IsInterfaceImplementation);
        }

        protected override string ResultDescription(Declaration declaration)
        {
            return string.Format(InspectionResults.UnderscoreInPublicClassModuleMemberInspection, declaration.IdentifierName);
        }
    }
}
