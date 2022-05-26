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
    /// Locates unqualified Worksheet.Range/Cells/Columns/Rows member calls inside worksheet modules, that implicitly refer to the containing sheet component.
    /// </summary>
    /// <hostApp name="Excel" />
    /// <why>
    /// Implicit references inside a worksheet document module can easily be mistaken for implicit references to the active worksheet (ActiveSheet), which is the behavior in all other module types.
    /// By explicitly qualifying these member calls with 'Me', the ambiguity can be resolved. If the intent is to refer to the active worksheet, qualify with 'ActiveSheet' to prevent a bug.
    /// </why>
    /// <example hasResult="true">
    /// <module name="Sheet1" type="Document Module">
    /// <![CDATA[
    /// Private Sub Example()
    ///     Dim foo As Range
    ///     Set foo = Range("A1") ' Worksheet.Range implicitly from containing worksheet
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="Sheet1" type="Document Module">
    /// <![CDATA[
    /// Private Sub Example()
    ///     Dim foo As Range
    ///     Set foo = Me.Range("A1")
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    [RequiredHost("Excel")]
    internal sealed class ImplicitContainingWorksheetReferenceInspection : ImplicitSheetReferenceInspectionBase
    {
        public ImplicitContainingWorksheetReferenceInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {}

        protected override bool IsResultReference(IdentifierReference reference, DeclarationFinder finder)
        {
            return Declaration.GetModuleParent(reference.ParentNonScoping) is DocumentModuleDeclaration document
                   && document.SupertypeNames.Contains("Worksheet")
                   && reference.QualifyingReference == null; // if it's qualified, it's not an implicit reference
        }

        protected override string ResultDescription(IdentifierReference reference)
        {
            return string.Format(
                InspectionResults.ImplicitContainingWorksheetReferenceInspection,
                reference.Declaration.IdentifierName);
        }
    }
}