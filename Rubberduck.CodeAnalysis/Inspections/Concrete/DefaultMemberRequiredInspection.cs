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
    /// Locates indexed default member calls for which the corresponding object does not have a suitable suitable default member. 
    /// </summary>
    /// <why>
    /// The VBA compiler does not check whether the necessary default member is present. Instead there is a runtime error whenever the runtime type fails to have the default member.
    /// </why>
    /// <example hasResult="true">
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
    /// <example hasResult="false">
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
    public class DefaultMemberRequiredInspection : InspectionBase
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public DefaultMemberRequiredInspection(RubberduckParserState state)
            : base(state)
        {
            _declarationFinderProvider = state;

            //This will most likely cause a runtime error. The exceptions are rare and should be refactored or made explicit with an @Ignore annotation.
            Severity = CodeInspectionSeverity.Error;
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var finder = _declarationFinderProvider.DeclarationFinder;

            var failedIndexedDefaultMemberAccesses = finder.FailedIndexedDefaultMemberAccesses();
            return failedIndexedDefaultMemberAccesses
                .Where(failedIndexedDefaultMemberAccess => !IsIgnored(failedIndexedDefaultMemberAccess))
                .Select(failedIndexedDefaultMemberAccess => InspectionResult(failedIndexedDefaultMemberAccess, _declarationFinderProvider));
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

        private string ResultDescription(IdentifierReference failedIndexedDefaultMemberAccess)
        {
            var expression = failedIndexedDefaultMemberAccess.IdentifierName;
            var typeName = failedIndexedDefaultMemberAccess.Declaration?.FullAsTypeName;
            return string.Format(InspectionResults.DefaultMemberRequiredInspection, expression, typeName);
        }
    }
}