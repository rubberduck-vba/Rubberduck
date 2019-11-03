using System.Collections.Generic;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Inspections.Extensions;
using Rubberduck.Parsing.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
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
    /// <![CDATA[
    /// Class1:
    ///
    /// Public Function Foo(index As Long) As Long
    /// 'No default member attribute
    /// End Function
    ///
    /// ------------------------------
    /// Module1:
    /// 
    /// Public Sub DoIt()
    ///     Dim cls As Class1
    ///     Dim bar As Variant
    ///     Set cls = New Class1
    ///     bar = cls(0) 
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasresult="false">
    /// <![CDATA[
    /// Class1:
    ///
    /// Public Function Foo(index As Long) As Long
    /// Attribute Foo.UserMemId = 0
    /// End Function
    ///
    /// ------------------------------
    /// Module1:
    /// 
    /// Public Sub DoIt()
    ///     Dim cls As Class1
    ///     Dim bar As Variant
    ///     Set cls = New Class1
    ///     bar = cls(0) 
    /// End Sub
    /// ]]>
    /// </example>
    public class DefaultMemberRequiredInspection : IdentifierReferenceInspectionBase
    {
        public DefaultMemberRequiredInspection(RubberduckParserState state)
            : base(state)
        {
            //This will most likely cause a runtime error. The exceptions are rare and should be refactored or made explicit with an @Ignore annotation.
            Severity = CodeInspectionSeverity.Error;
        }

        protected override IEnumerable<IdentifierReference> ReferencesInModule(QualifiedModuleName module)
        {
            return DeclarationFinderProvider.DeclarationFinder.FailedIndexedDefaultMemberAccesses(module);
        }

        protected override bool IsResultReference(IdentifierReference failedIndexedDefaultMemberAccess)
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