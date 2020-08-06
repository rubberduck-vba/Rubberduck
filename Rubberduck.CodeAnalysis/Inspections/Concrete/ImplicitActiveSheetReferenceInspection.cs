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
    /// Locates unqualified Worksheet.Range/Cells/Columns/Rows member calls that implicitly refer to ActiveSheet.
    /// </summary>
    /// <reference name="Excel" />
    /// <why>
    /// Implicit references to the active worksheet rarely mean to be working with *whatever worksheet is currently active*. 
    /// By explicitly qualifying these member calls with a specific Worksheet object, the assumptions are removed, the code
    /// is more robust, and will be less likely to throw run-time error 1004 or produce unexpected results
    /// when the active sheet isn't the expected one.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Private Sub Example()
    ///     Dim foo As Range
    ///     Set foo = Sheet1.Range(Cells(1, 1), Cells(1, 10)) ' Worksheet.Cells implicitly from ActiveSheet; error 1004 if that isn't Sheet1.
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Private Sub Example()
    ///     Dim foo As Range
    ///     With Sheet1
    ///         Set foo = .Range(.Cells(1, 1), .Cells(1, 10)) ' all member calls are made against the With block object
    ///     End With
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    [RequiredLibrary("Excel")]
    internal sealed class ImplicitActiveSheetReferenceInspection : IdentifierReferenceInspectionFromDeclarationsBase
    {
        public ImplicitActiveSheetReferenceInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {}

        protected override IEnumerable<Declaration> ObjectionableDeclarations(DeclarationFinder finder)
        {
            var excel = finder.Projects
                .SingleOrDefault(item => !item.IsUserDefined
                                         && item.IdentifierName == "Excel");
            if (excel == null)
            {
                return Enumerable.Empty<Declaration>();
            }

            var globalModules = GlobalObjectClassNames
                .Select(className => finder.FindClassModule(className, excel, true))
                .OfType<ModuleDeclaration>();

            return globalModules
                .SelectMany(moduleClass => moduleClass.Members)
                .Where(declaration => TargetMemberNames.Contains(declaration.IdentifierName) 
                                      && declaration.DeclarationType.HasFlag(DeclarationType.Member)
                                      && declaration.AsTypeName == "Range");
        }

        private static readonly string[] GlobalObjectClassNames =
        {
            "Global", "_Global"
        };

        private static readonly string[] TargetMemberNames =
        {
            "Cells", "Range", "Columns", "Rows"
        };

        protected override string ResultDescription(IdentifierReference reference)
        {
            return string.Format(
                InspectionResults.ImplicitActiveSheetReferenceInspection,
                reference.Declaration.IdentifierName);
        }
    }
}
