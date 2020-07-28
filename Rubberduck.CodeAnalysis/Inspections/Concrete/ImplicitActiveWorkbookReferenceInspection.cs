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
    internal sealed class ImplicitActiveWorkbookReferenceInspection : IdentifierReferenceInspectionFromDeclarationsBase
    {
        public ImplicitActiveWorkbookReferenceInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {}

        private static readonly string[] InterestingMembers =
        {
            "Worksheets", "Sheets", "Names"
        };

        private static readonly string[] InterestingClasses =
        {
            "_Global", "_Application", "Global", "Application"
        };

        protected override IEnumerable<Declaration> ObjectionableDeclarations(DeclarationFinder finder)
        {
            var excel = finder.Projects
                .SingleOrDefault(project => project.IdentifierName == "Excel" && !project.IsUserDefined);
            if (excel == null)
            {
                return Enumerable.Empty<Declaration>();
            }

            var relevantClasses = InterestingClasses
                .Select(className => finder.FindClassModule(className, excel, true))
                .OfType<ModuleDeclaration>();

            var relevantProperties = relevantClasses
                .SelectMany(classDeclaration => classDeclaration.Members)
                .OfType<PropertyGetDeclaration>()
                .Where(member => InterestingMembers.Contains(member.IdentifierName));

            return relevantProperties;
        }

        protected override string ResultDescription(IdentifierReference reference)
        {
            var referenceText = reference.Context.GetText();
            return string.Format(InspectionResults.ImplicitActiveWorkbookReferenceInspection, referenceText);
        }
    }
}
