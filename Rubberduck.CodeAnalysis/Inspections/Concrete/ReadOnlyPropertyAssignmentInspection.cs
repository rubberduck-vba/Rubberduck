using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;
using System.Linq;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Identifies Property assigment references where Set or Let Property Members do not exist.
    /// </summary>
    /// <why>
    /// In general, the VBE editor catches this type of error and will not compile.  However, there are 
    /// a few scenarios where the error is overlooked by the compiler and an error is generated at runtime.  
    /// To avoid the runtime error scenarios, the inspection flags all assignment references of a read-only property.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyDataObject" type="Class Module">
    /// <![CDATA[
    /// Public myData As Long
    /// ]]>
    /// </module>
    /// <module name="Client" type="Standard Module">
    /// <![CDATA[
    /// Private myDataObj As MyDataObject
    /// 
    /// Public Sub Test()
    ///     Set TheData = new MyDataObject
    /// End Sub
    /// 
    /// Public Property Get TheData() As MyDataObject
    ///     Set TheData = myDataObj
    /// End Property
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyDataObject" type="Class Module">
    /// <![CDATA[
    /// Public myData As Long
    /// ]]>
    /// </module>
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Private myDataObj As MyDataObject
    /// 
    /// Public Sub Test()
    ///     Set TheData = new MyDataObject
    /// End Sub
    /// 
    /// Public Property Get TheData() As MyDataObject
    ///     Set TheData = myDataObj
    /// End Property
    /// Public Property Set TheData(RHS As MyDataObject)
    ///     Set myDataObj = RHS
    /// End Property
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Private myData As Variant
    /// 
    /// Public Sub Test()
    ///     TheData = 45
    /// End Sub
    /// 
    /// Public Property Get TheData() As Variant
    ///     TheData = myData
    /// End Property
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="Client" type="Standard Module">
    /// <![CDATA[
    /// Private myData As Variant
    /// 
    /// Public Sub Test()
    ///     TheData = 45
    /// End Sub
    /// 
    /// Public Property Get TheData() As Variant
    ///     TheData = myData
    /// End Property
    /// Public Property Let TheData(RHS As Variant)
    ///     myData = RHS
    /// End Property
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class ReadOnlyPropertyAssignmentInspection : IdentifierReferenceInspectionBase
    {
        public ReadOnlyPropertyAssignmentInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        { }

        protected override bool IsResultReference(IdentifierReference reference, DeclarationFinder finder)
        {
            if (!reference.Declaration.DeclarationType.HasFlag(DeclarationType.Property))
            {
                return false;
            }

            //Ignore assignment expressions found within Property Get declaration contexts
            if (!IsReadOnlyPropertyReference(reference, finder)
                || reference.Declaration.Context.Contains(reference.Context))
            {
                return false;
            }

            return reference.IsAssignment;
        }

        private bool IsReadOnlyPropertyReference(IdentifierReference reference, DeclarationFinder finder)
        {
            var propertyDeclarations = finder.MatchName(reference.Declaration.IdentifierName)
                .Where(d => d.DeclarationType.HasFlag(DeclarationType.Property)
                    && d.QualifiedModuleName == reference.QualifiedModuleName);

            return propertyDeclarations.Count() == 1
                && propertyDeclarations.First().DeclarationType.HasFlag(DeclarationType.PropertyGet);
        }

        protected override string ResultDescription(IdentifierReference reference)
        {
            var identifierName = reference.IdentifierName;
            return string.Format(
                InspectionResults.ReadOnlyPropertyAssignmentInspection, identifierName);
        }
    }
}