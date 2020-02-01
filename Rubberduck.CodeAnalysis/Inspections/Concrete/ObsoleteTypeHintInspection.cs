using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Flags declarations where a type hint is used in place of an 'As' clause.
    /// </summary>
    /// <why>
    /// Type hints were made obsolete when declaration syntax introduced the 'As' keyword. Prefer explicit type names over type hint symbols.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim foo$
    ///     foo = "some string"
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim foo As String
    ///     foo = "some string"
    /// End Sub
    /// ]]>
    /// </example>
    public sealed class ObsoleteTypeHintInspection : InspectionBase
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public ObsoleteTypeHintInspection(RubberduckParserState state)
            : base(state)
        {
            _declarationFinderProvider = state;
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var finder = _declarationFinderProvider.DeclarationFinder;

            var results = new List<IInspectionResult>();
            foreach (var moduleDeclaration in finder.UserDeclarations(DeclarationType.Module))
            {
                if (moduleDeclaration == null)
                {
                    continue;
                }

                var module = moduleDeclaration.QualifiedModuleName;
                results.AddRange(DoGetInspectionResults(module, finder));
            }

            return results;
        }

        private  IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module)
        {
            var finder = _declarationFinderProvider.DeclarationFinder;
            return DoGetInspectionResults(module, finder);
        }

        private IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module, DeclarationFinder finder)
        {
            var declarationResults = DeclarationResults(module, finder);
            var referenceResults = ReferenceResults(module, finder);
            return declarationResults
                .Concat(referenceResults);
        }

        private IEnumerable<IInspectionResult> DeclarationResults(QualifiedModuleName module, DeclarationFinder finder)
        {
            var objectionableDeclarations = finder.Members(module)
                .Where(declaration => declaration.HasTypeHint);
            return objectionableDeclarations.Select(InspectionResult);
        }

        private IInspectionResult InspectionResult(Declaration declaration)
        {
            return new DeclarationInspectionResult(
                this,
                ResultDescription(declaration),
                declaration);
        }

        private static string ResultDescription(Declaration declaration)
        {
            var declarationTypeName = declaration.DeclarationType.ToString().ToLower();
            var identifierName = declaration.IdentifierName;
            return string.Format(
                InspectionResults.ObsoleteTypeHintInspection,
                InspectionsUI.Inspections_Declaration,
                declarationTypeName,
                identifierName);
        }

        private IEnumerable<IInspectionResult> ReferenceResults(QualifiedModuleName module, DeclarationFinder finder)
        {
            var objectionableReferences = finder.IdentifierReferences(module)
                .Where(reference => reference?.Declaration != null
                                    && reference.Declaration.IsUserDefined
                                    && reference.HasTypeHint());
            return objectionableReferences
                .Select(reference => InspectionResult(reference, _declarationFinderProvider));
        }

        private IInspectionResult InspectionResult(IdentifierReference reference, IDeclarationFinderProvider declarationFinderProvider)
        {
            return new IdentifierReferenceInspectionResult(
                this,
                ResultDescription(reference),
                declarationFinderProvider,
                reference);
        }

        private string ResultDescription(IdentifierReference reference)
        {
            var declarationTypeName = reference.Declaration.DeclarationType.ToString().ToLower();
            var identifierName = reference.IdentifierName;
            return string.Format(InspectionResults.ObsoleteTypeHintInspection,
                InspectionsUI.Inspections_Usage,
                declarationTypeName,
                identifierName);
        }
    }
}
