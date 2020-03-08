using System.Collections.Generic;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.CodeAnalysis.Inspections.Results;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;
using Rubberduck.VBEditor;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Flags declarations where a type hint is used in place of an 'As' clause.
    /// </summary>
    /// <why>
    /// Type hints were made obsolete when declaration syntax introduced the 'As' keyword. Prefer explicit type names over type hint symbols.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim foo$
    ///     foo = "some string"
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim foo As String
    ///     foo = "some string"
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class ObsoleteTypeHintInspection : InspectionBase
    {
        public ObsoleteTypeHintInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {}

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults(DeclarationFinder finder)
        {
            return finder.UserDeclarations(DeclarationType.Module)
                .Where(module => module != null)
                .SelectMany(module => DoGetInspectionResults(module.QualifiedModuleName, finder));
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module, DeclarationFinder finder)
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
                .Select(reference => InspectionResult(reference, finder));
        }

        private IInspectionResult InspectionResult(IdentifierReference reference, DeclarationFinder finder)
        {
            return new IdentifierReferenceInspectionResult(
                this,
                ResultDescription(reference),
                finder,
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
