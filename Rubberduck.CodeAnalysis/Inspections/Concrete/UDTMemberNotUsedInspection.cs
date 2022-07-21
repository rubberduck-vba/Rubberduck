using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.CodeAnalysis.Inspections.Extensions;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;
using System.Linq;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Warns about User Defined Type (UDT) members that are never referenced.
    /// </summary>
    /// <why>
    /// Declarations that are never used should be removed.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Private Type TTestModule
    ///     FirstVal As Long
    /// End Type
    /// 
    /// Private this As TTestModule
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Private Type TTestModule
    ///     FirstVal As Long
    /// End Type
    /// 
    /// Private this As TTestModule
    /// 
    /// 'UDT Member is assigned but not used
    /// Public Sub DoSomething()
    ///     this.FirstVal = 42
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Private Type TTestModule
    ///     FirstVal As Long
    /// End Type
    /// 
    /// Private this As TTestModule
    /// 
    /// 'UDT Member is assigned and read
    /// Public Sub DoSomething()
    ///     this.FirstVal = 42
    ///     Debug.Print this.FirstVal
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class UDTMemberNotUsedInspection : DeclarationInspectionBase
    {
        public UDTMemberNotUsedInspection(IDeclarationFinderProvider declarationFinderProvider)
                : base(declarationFinderProvider, DeclarationType.UserDefinedTypeMember)
        {}

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder)
        {
            return declaration.DeclarationType.Equals(DeclarationType.UserDefinedTypeMember)
                && !declaration.References.Any();
        }

        protected override string ResultDescription(Declaration declaration)
        {
            var declarationType = declaration.DeclarationType.ToLocalizedString();
            var declarationName = declaration.IdentifierName;
            return string.Format(
                InspectionResults.IdentifierNotUsedInspection,
                declarationType,
                declarationName);
        }
    }
}
