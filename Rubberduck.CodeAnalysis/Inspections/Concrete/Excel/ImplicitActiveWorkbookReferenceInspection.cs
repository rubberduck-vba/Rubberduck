using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.Inspections.Inspections.Extensions;

namespace Rubberduck.Inspections.Concrete
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
    /// <example hasResults="true">
    /// <![CDATA[
    /// Private Sub Example()
    ///     Dim summarySheet As Worksheet
    ///     Set summarySheet = Worksheets("Summary") ' unqualified Worksheets is implicitly querying the active workbook's Worksheets collection.
    /// End Sub
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Private Sub Example(ByVal book As Workbook)
    ///     Dim summarySheet As Worksheet
    ///     Set summarySheet = book.Worksheets("Summary")
    /// End Sub
    /// ]]>
    /// </example>
    [RequiredLibrary("Excel")]
    public sealed class ImplicitActiveWorkbookReferenceInspection : InspectionBase
    {
        public ImplicitActiveWorkbookReferenceInspection(RubberduckParserState state)
            : base(state) { }

        private static readonly string[] InterestingMembers =
        {
            "Worksheets", "Sheets", "Names"
        };

        private static readonly string[] InterestingClasses =
        {
            "_Global", "_Application", "Global", "Application"
        };

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var excel = State.DeclarationFinder.Projects.SingleOrDefault(item => !item.IsUserDefined && item.IdentifierName == "Excel");
            if (excel == null)
            {
                return Enumerable.Empty<IInspectionResult>();
            }

            var targetProperties = BuiltInDeclarations
                .OfType<PropertyGetDeclaration>()
                .Where(x => InterestingMembers.Contains(x.IdentifierName) && InterestingClasses.Contains(x.ParentDeclaration?.IdentifierName))
                .ToList();

            // only inspects references, must filter ignores manually, because default filtering doesn't work here
            var members = targetProperties.SelectMany(item =>
                item.References.Where(reference => !reference.IsIgnoringInspectionResultFor(AnnotationName)));

            return members.Select(issue => new IdentifierReferenceInspectionResult(this,
                                                                string.Format(InspectionResults.ImplicitActiveWorkbookReferenceInspection, issue.Context.GetText()),
                                                                State,
                                                                issue));
        }
    }
}
