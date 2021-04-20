using System.Collections.Generic;
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

        private IReadOnlyList<Declaration> _applicationCandidates;

        protected override bool IsResultReference(IdentifierReference reference, DeclarationFinder finder)
        {
            var qualifiers = base.GetQualifierCandidates(reference, finder);
            var isQualified = qualifiers.Any();
            var document = Declaration.GetModuleParent(reference.ParentNonScoping) as DocumentModuleDeclaration;

            var isHostWorkbook = (document?.SupertypeNames.Contains("Workbook") ?? false)
                && (document?.ProjectId?.Equals(reference.QualifiedModuleName.ProjectId) ?? false);

            if (!isQualified)
            {
                // unqualified calls aren't referring to ActiveWorkbook only inside a Workbook module:
                return !isHostWorkbook;
            }
            else
            {
                if (_applicationCandidates == null)
                {
                    var applicationClass = finder.FindClassModule("Application", base.Excel, includeBuiltIn: true);
                    // note: underscored declarations would be for unqualified calls
                    var workbookClass = finder.FindClassModule("Workbook", base.Excel, includeBuiltIn: true);
                    var worksheetClass = finder.FindClassModule("Worksheet", base.Excel, includeBuiltIn: true);
                    var hostBook = finder.UserDeclarations(DeclarationType.Document)
                        .Cast<DocumentModuleDeclaration>()
                        .SingleOrDefault(doc => doc.ProjectId.Equals(reference.QualifiedModuleName.ProjectId)
                            && doc.SupertypeNames.Contains("Workbook"));

                    _applicationCandidates = finder.MatchName("Application")
                        .Where(m => m.Equals(applicationClass) 
                        || (m.ParentDeclaration.Equals(workbookClass) && m.DeclarationType.HasFlag(DeclarationType.PropertyGet))
                        || (m.ParentDeclaration.Equals(worksheetClass) && m.DeclarationType.HasFlag(DeclarationType.PropertyGet))
                        || (m.ParentDeclaration.Equals(hostBook) && m.DeclarationType.HasFlag(DeclarationType.PropertyGet)))
                        .ToList();
                }

                // qualified calls are referring to ActiveWorkbook if qualifier is the Application object:
                return _applicationCandidates.Any(candidate => qualifiers.Any(q => q.Equals(candidate)));
            }
        }

        protected override string ResultDescription(IdentifierReference reference)
        {
            var referenceText = reference.Context.GetText();
            return string.Format(InspectionResults.ImplicitActiveWorkbookReferenceInspection, referenceText);
        }
    }
}
