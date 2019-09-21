using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Inspections.Inspections.Extensions;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Identifies the use of bang notation, formally known as dictionary access expression.
    /// </summary>
    /// <why>
    /// A dictionary access expression looks like a strongly typed call, but it actually is a stringly typed access to the parameterized default member of the object. 
    /// </why>
    /// <example hasResult="true">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal wkb As Excel.Workbook)
    ///     wkb.Worksheets!MySheet.Range("A1").Value = 42
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResult="true">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal wkb As Excel.Workbook)
    ///     With wkb.Worksheets
    ///         !MySheet.Range("A1").Value = 42
    ///     End With
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResult="false">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal wkb As Excel.Workbook)
    ///     wkb.Worksheets("MySheet").Range("A1").Value = 42
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResult="false">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal wkb As Excel.Workbook)
    ///     wkb.Worksheets.Item("MySheet").Range("A1").Value = 42
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResult="false">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal wkb As Excel.Workbook)
    ///     With wkb.Worksheets
    ///         .Item("MySheet").Range("A1").Value = 42
    ///     End With
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class UseOfBangNotationInspection : InspectionBase
    {
        public UseOfBangNotationInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var results = new List<IInspectionResult>();
            foreach (var moduleDeclaration in State.DeclarationFinder.UserDeclarations(DeclarationType.Module))
            {
                if (moduleDeclaration == null || moduleDeclaration.IsIgnoringInspectionResultFor(AnnotationName))
                {
                    continue;
                }

                var module = moduleDeclaration.QualifiedModuleName;
                results.AddRange(DoGetInspectionResults(module));
            }

            return results;
        }

        private IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module)
        {
            var usesOfBang = State.DeclarationFinder
                .IdentifierReferences(module)
                .Where(IsRelevantReference);

            return usesOfBang
                .Select(useOfBang => InspectionResult(useOfBang, State))
                .ToList();
        }

        private bool IsRelevantReference(IdentifierReference reference)
        {
            return reference.IsIndexedDefaultMemberAccess
                   && reference.DefaultMemberRecursionDepth == 1
                   && reference.Context is VBAParser.DictionaryAccessContext
                   && !reference.IsIgnoringInspectionResultFor(AnnotationName);
        }

        private IInspectionResult InspectionResult(IdentifierReference dictionaryAccess, IDeclarationFinderProvider declarationFinderProvider)
        {
            return new IdentifierReferenceInspectionResult(this,
                ResultDescription(dictionaryAccess),
                declarationFinderProvider,
                dictionaryAccess);
        }

        private string ResultDescription(IdentifierReference dictionaryAccess)
        {
            var expression = dictionaryAccess.IdentifierName;
            return string.Format(InspectionResults.UseOfBangNotationInspection, expression);
        }
    }
}