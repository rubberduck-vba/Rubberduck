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
    /// Locates places in which a procedure needs to be called but an object variables has been provided that does not have a suitable default member. 
    /// </summary>
    /// <why>
    /// The VBA compiler does not check whether the necessary default member is present. Instead there is a runtime error whenever the runtime type fails to have the default member.
    /// </why>
    /// <example hasResult="true">
    /// <![CDATA[
    /// Class1:
    ///
    /// Public Sub Foo()
    /// 'No default member attribute
    /// End Sub
    ///
    /// ------------------------------
    /// Module1:
    /// 
    /// Public Sub DoIt()
    ///     Dim cls As Class1
    ///     Set cls = New Class1
    ///     cls 
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResult="false">
    /// <![CDATA[
    /// Class1:
    ///
    /// Public Sub Foo()
    /// Attribute Foo.UserMemId = 0
    /// End Sub
    ///
    /// ------------------------------
    /// Module1:
    /// 
    /// Public Sub DoIt()
    ///     Dim cls As Class1
    ///     Set cls = New Class1
    ///     cls 
    /// End Sub
    /// ]]>
    /// </example>
    public class ProcedureRequiredInspection : InspectionBase
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public ProcedureRequiredInspection(RubberduckParserState state)
            : base(state)
        {
            _declarationFinderProvider = state;

            //This will most likely cause a runtime error. The exceptions are rare and should be refactored or made explicit with an @Ignore annotation.
            Severity = CodeInspectionSeverity.Error;
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var finder = _declarationFinderProvider.DeclarationFinder;

            var failedProcedureCoercions = finder.FailedProcedureCoercions();
            return failedProcedureCoercions
                .Where(failedCoercion => !IsIgnored(failedCoercion))
                .Select(failedCoercion => InspectionResult(failedCoercion, _declarationFinderProvider));
        }

        private bool IsIgnored(IdentifierReference assignment)
        {
            return assignment.IsIgnoringInspectionResultFor(AnnotationName);
        }

        private IInspectionResult InspectionResult(IdentifierReference failedCoercion, IDeclarationFinderProvider declarationFinderProvider)
        {
            return new IdentifierReferenceInspectionResult(this,
                ResultDescription(failedCoercion),
                declarationFinderProvider,
                failedCoercion);
        }

        private string ResultDescription(IdentifierReference failedCoercion)
        {
            var expression = failedCoercion.IdentifierName;
            var typeName = failedCoercion.Declaration?.FullAsTypeName;
            return string.Format(InspectionResults.ProcedureRequiredInspection, expression, typeName);
        }
    }
}