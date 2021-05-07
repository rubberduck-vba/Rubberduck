using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.CodeAnalysis.Inspections.Attributes;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Locates unqualified Workbook.Worksheets/Sheets/Names member calls that implicitly refer to ActiveWorkbook.
    /// </summary>
    /// <reference name="Excel" />
    /// <why>
    /// Implicit references to the active workbook rarely mean to be working with *whatever workbook is currently active*. 
    /// By explicitly qualifying these member calls with a specific Workbook object, the assumptions are removed, the code
    /// is more robust, and will be less likely to throw run-time error 1004 or produce unexpected results
    /// when the active workbook isn't the expected one.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Private Sub Example()
    ///     Dim summarySheet As Worksheet
    ///     Set summarySheet = Worksheets("Summary") ' unqualified Worksheets is implicitly querying the active workbook's Worksheets collection.
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Private Sub Example(ByVal book As Workbook)
    ///     Dim summarySheet As Worksheet
    ///     Set summarySheet = book.Worksheets("Summary")
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    [RequiredLibrary("Excel")]
    internal sealed class ImplicitActiveWorkbookReferenceInspection : ImplicitWorkbookReferenceInspectionBase
    {
        public ImplicitActiveWorkbookReferenceInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider) { }

        protected override bool IsResultReference(IdentifierReference reference, DeclarationFinder finder)
        {
            var isQualified = reference.QualifyingReference != null;
            if (!isQualified)
            {
                var document = Declaration.GetModuleParent(reference.ParentNonScoping) as DocumentModuleDeclaration;

                var isHostWorkbook = (document?.SupertypeNames.Contains("Workbook") ?? false)
                                     && (document?.ProjectId?.Equals(reference.QualifiedModuleName.ProjectId) ?? false);

                // unqualified calls aren't referring to ActiveWorkbook only inside a Workbook module:
                return !isHostWorkbook;
            }

            if (reference.QualifyingReference.Declaration == null)
            {
                //This should really only happen on unbound member calls and then the current reference would also be unbound.
                //So, if we end up here, we have no idea and bail out.
                return false;
            }

            var excelProjectId = Excel(finder).ProjectId;
            var applicationCandidates = finder.MatchName("Application")
                .Where(m =>  m.ProjectId.Equals(excelProjectId) 
                             && ( m.DeclarationType == DeclarationType.PropertyGet 
                                || m.DeclarationType == DeclarationType.ClassModule));

            var qualifyingDeclaration = reference.QualifyingReference.Declaration;

            // qualified calls are referring to ActiveWorkbook if qualifier is the Application object:
            return applicationCandidates.Any(candidate => qualifyingDeclaration.Equals(candidate));
        }

        protected override string ResultDescription(IdentifierReference reference)
        {
            var referenceText = reference.Context.GetText();
            return string.Format(InspectionResults.ImplicitActiveWorkbookReferenceInspection, referenceText);
        }
    }
}
