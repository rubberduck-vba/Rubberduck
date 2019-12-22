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
    /// Locates places in which a value needs to be accessed but an object variables has been provided that does not have a suitable default member. 
    /// </summary>
    /// <why>
    /// The VBA compiler does not check whether the necessary default member is present. Instead there is a runtime error whenever the runtime type fails to have the default member.
    /// </why>
    /// <example hasresult="true">
    /// <![CDATA[
    /// Class1:
    ///
    /// Public Function Foo() As Long
    /// 'No default member attribute
    /// End Function
    ///
    /// ------------------------------
    /// Module1:
    /// 
    /// Public Sub DoIt()
    ///     Dim cls As Class1
    ///     Dim bar As Variant
    ///
    ///     Set cls = New Class1
    ///     bar = cls + 42 
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasresult="false">
    /// <![CDATA[
    /// Class1:
    ///
    /// Public Function Foo() As Long
    /// Attribute Foo.UserMemId = 0
    /// End Function
    ///
    /// ------------------------------
    /// Module1:
    /// 
    /// Public Sub DoIt()
    ///     Dim cls As Class1
    ///     Dim bar As Variant
    ///
    ///     Set cls = New Class1
    ///     bar = cls + 42 
    /// End Sub
    /// ]]>
    /// </example>
    public class ValueRequiredInspection : IdentifierReferenceInspectionBase
    {
        public ValueRequiredInspection(RubberduckParserState state)
            : base(state)
        {
            //This will most likely cause a runtime error. The exceptions are rare and should be refactored or made explicit with an @Ignore annotation.
            Severity = CodeInspectionSeverity.Error;
        }

        protected override IEnumerable<IdentifierReference> ReferencesInModule(QualifiedModuleName module)
        {
            return DeclarationFinderProvider.DeclarationFinder.FailedLetCoercions(module);
        }

        protected override bool IsResultReference(IdentifierReference failedLetCoercion)
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
