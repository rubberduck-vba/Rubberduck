using System.Collections.Generic;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;
using Rubberduck.VBEditor;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Locates places in which a value needs to be accessed but an object variables has been provided that does not have a suitable default member. 
    /// </summary>
    /// <why>
    /// The VBA compiler does not check whether the necessary default member is present. Instead there is a runtime error whenever the runtime type fails to have the default member.
    /// </why>
    /// <example hasresult="true">
    /// <module name="Class1" type="Class Module">
    /// <![CDATA[
    /// Public Function Foo() As Long
    /// 'No default member attribute
    /// End Function
    /// ]]>
    /// </module>
    /// <module name="Module1" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoIt()
    ///     Dim cls As Class1
    ///     Dim bar As Variant
    ///
    ///     Set cls = New Class1
    ///     bar = cls + 42 
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasresult="false">
    /// <module name="Class1" type="Class Module">
    /// <![CDATA[
    /// Public Function Foo() As Long
    /// Attribute Foo.UserMemId = 0
    /// End Function
    /// ]]>
    /// </module>
    /// <module name="Module1" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoIt()
    ///     Dim cls As Class1
    ///     Dim bar As Variant
    ///
    ///     Set cls = New Class1
    ///     bar = cls + 42 
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal class ValueRequiredInspection : IdentifierReferenceInspectionBase
    {
        public ValueRequiredInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            //This will most likely cause a runtime error. The exceptions are rare and should be refactored or made explicit with an @Ignore annotation.
            Severity = CodeInspectionSeverity.Error;
        }

        protected override IEnumerable<IdentifierReference> ReferencesInModule(QualifiedModuleName module, DeclarationFinder finder)
        {
            return finder.FailedLetCoercions(module);
        }

        protected override bool IsResultReference(IdentifierReference failedLetCoercion, DeclarationFinder finder)
        {
            return !failedLetCoercion.IsAssignment;
        }

        protected override string ResultDescription(IdentifierReference failedLetCoercion)
        {
            var expression = failedLetCoercion.IdentifierName;
            var typeName = failedLetCoercion.Declaration?.FullAsTypeName;
            return string.Format(InspectionResults.ValueRequiredInspection, expression, typeName);
        }
    }
}
