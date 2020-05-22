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
    /// Locates indexed default member calls for which the corresponding object does not have a suitable suitable default member. 
    /// </summary>
    /// <why>
    /// The VBA compiler does not check whether the necessary default member is present. Instead there is a runtime error whenever the runtime type fails to have the default member.
    /// </why>
    /// <example hasresult="true">
    /// <module name="Class1" type="Class Module">
    /// <![CDATA[
    /// Public Function Foo(index As Long) As Long
    /// 'No default member attribute
    /// End Function
    /// ]]>
    /// </module>
    /// <module name="Module1" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoIt()
    ///     Dim cls As Class1
    ///     Dim bar As Variant
    ///     Set cls = New Class1
    ///     bar = cls(0) 
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasresult="false">
    /// <module name="Class1" type="Class Module">
    /// <![CDATA[
    /// Public Function Foo(index As Long) As Long
    /// Attribute Foo.UserMemId = 0
    /// End Function
    /// ]]>
    /// </module>
    /// <module name="Module1" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoIt()
    ///     Dim cls As Class1
    ///     Dim bar As Variant
    ///     Set cls = New Class1
    ///     bar = cls(0) 
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal class DefaultMemberRequiredInspection : IdentifierReferenceInspectionBase
    {
        public DefaultMemberRequiredInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            //This will most likely cause a runtime error. The exceptions are rare and should be refactored or made explicit with an @Ignore annotation.
            Severity = CodeInspectionSeverity.Error;
        }

        protected override IEnumerable<IdentifierReference> ReferencesInModule(QualifiedModuleName module, DeclarationFinder finder)
        {
            return finder.FailedIndexedDefaultMemberAccesses(module);
        }

        protected override bool IsResultReference(IdentifierReference failedIndexedDefaultMemberAccess, DeclarationFinder finder)
        {
            return true;
        }

        protected override string ResultDescription(IdentifierReference failedIndexedDefaultMemberAccess)
        {
            var expression = failedIndexedDefaultMemberAccess.IdentifierName;
            var typeName = failedIndexedDefaultMemberAccess.Declaration?.FullAsTypeName;
            return string.Format(InspectionResults.DefaultMemberRequiredInspection, expression, typeName);
        }
    }
}