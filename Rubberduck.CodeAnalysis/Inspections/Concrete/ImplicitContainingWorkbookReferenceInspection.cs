using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.CodeAnalysis.Inspections.Attributes;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Locates unqualified Workbook.Worksheets/Sheets/Names member calls inside workbook document modules that implicitly refer to the containing workbook.
    /// </summary>
    /// <reference name="Excel" />
    /// <why>
    /// Implicit references inside a workbook document module can be mistakes for implicit references to the active workbook, which is the behavior in all other modules 
    /// By explicitly qualifying these member calls with Me, the ambiguity can be resolved.
    /// </why>
    /// <example hasResult="true">
    /// <module name="ThisWorkbook" type="Document Module">
    /// <![CDATA[
    /// Private Sub Example()
    ///     Dim summarySheet As Worksheet
    ///     Set summarySheet = Worksheets("Summary") ' unqualified Worksheets is implicitly querying the containing workbook's Worksheets collection.
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="ThisWorkbook" type="Document Module">
    /// <![CDATA[
    /// Private Sub Example()
    ///     Dim summarySheet As Worksheet
    ///     Set summarySheet = Me.Worksheets("Summary")
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    [RequiredLibrary("Excel")]
    internal sealed class ImplicitContainingWorkbookReferenceInspection : ImplicitWorkbookReferenceInspectionBase
    {
        public ImplicitContainingWorkbookReferenceInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        { }

        private static readonly string[] ParentsNeverReferringToContainingWorkbook = new[]
        {
            "Application", "_Application"
        };

        protected override IEnumerable<Declaration> ObjectionableDeclarations(DeclarationFinder finder)
        {
            return base.ObjectionableDeclarations(finder)
                .Where(declaration => !ParentsNeverReferringToContainingWorkbook
                    .Any(name => declaration.ParentDeclaration.IdentifierName.Equals(name, System.StringComparison.InvariantCultureIgnoreCase)));
        }

        protected override bool IsResultReference(IdentifierReference reference, DeclarationFinder finder)
        {
            var qualifiers = base.GetQualifierCandidates(reference, finder);
            return Declaration.GetModuleParent(reference.ParentScoping) is DocumentModuleDeclaration document
                && document.SupertypeNames.Contains("Workbook")
                && (qualifiers.Any());
        }

        protected override string ResultDescription(IdentifierReference reference)
        {
            var referenceText = reference.Context.GetText();
            return string.Format(
                InspectionResults.ImplicitContainingWorkbookReferenceInspection, 
                referenceText);
        }
    }
}