using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Inspections.Extensions;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Locates places in which a value needs to be accessed but an object variables has been provided that does not have a suitable default member. 
    /// </summary>
    /// <why>
    /// The VBA compiler does not check whether the necessary default member is present. Instead there is a runtime error whenever the runtime type fails to have the default member.
    /// </why>
    /// <example hasResult="true">
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
    /// <example hasResult="false">
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
    public class ValueRequiredInspection : InspectionBase
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public ValueRequiredInspection(RubberduckParserState state)
            : base(state)
        {
            _declarationFinderProvider = state;

            //This will most likely cause a runtime error. The exceptions are rare and should be refactored or made explicit with an @Ignore annotation.
            Severity = CodeInspectionSeverity.Error;
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var finder = _declarationFinderProvider.DeclarationFinder;

            //Assignments are already covered by the ObjectVariableNotSetInspection.
            var failedLetCoercionAccesses = finder.FailedLetCoercions()
                .Where(failedLetCoercion => !failedLetCoercion.IsAssignment);

            return failedLetCoercionAccesses
                .Where(failedLetCoercion => !IsIgnored(failedLetCoercion))
                .Select(failedLetCoercion => InspectionResult(failedLetCoercion, _declarationFinderProvider));
        }

        private bool IsIgnored(IdentifierReference assignment)
        {
            return assignment.IsIgnoringInspectionResultFor(AnnotationName);
        }

        private IInspectionResult InspectionResult(IdentifierReference failedLetCoercion, IDeclarationFinderProvider declarationFinderProvider)
        {
            return new IdentifierReferenceInspectionResult(this,
                ResultDescription(failedLetCoercion),
                declarationFinderProvider,
                failedLetCoercion);
        }

        private string ResultDescription(IdentifierReference failedLetCoercion)
        {
            var expression = failedLetCoercion.IdentifierName;
            var typeName = failedLetCoercion.Declaration?.FullAsTypeName;
            return string.Format(InspectionResults.ValueRequiredInspection, expression, typeName);
        }
    }
}
