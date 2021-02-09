using System.Collections.Generic;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.CodeAnalysis.Inspections.Attributes;
using Rubberduck.InternalApi.Extensions;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Warns about late-bound WorksheetFunction calls made against the extended interface of the Application object.
    /// </summary>
    /// <reference name="Excel" />
    /// <why>
    /// An early-bound, equivalent function exists in the object returned by the Application.WorksheetFunction property; 
    /// late-bound member calls will fail at run-time with error 438 if there is a typo (a typo fails to compile for an early-bound member call); 
    /// given invalid inputs, these late-bound member calls return a Variant/Error value that cannot be coerced into another type.
    /// The equivalent early-bound member calls raise a more VB-idiomatic, trappable runtime error given the same invalid inputs: 
    /// trying to compare or assign a Variant/Error to another data type will throw error 13 "type mismatch" at run-time. 
    /// A Variant/Error value cannot be coerced into any other data type, be it for assignment or comparison.
    /// 
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Private Sub Example()
    ///     Debug.Print Application.Sum(Array(1, 2, 3), 4, 5, "ABC") ' outputs "Error 2015" (no run-time error is raised).
    /// 
    ///     Dim foo As Long
    ///     foo = Application.Sum(Array(1, 2, 3), 4, 5, "ABC") ' error 13 "type mismatch". Variant/Error can't be coerced to Long.
    /// 
    ///     If Application.Sum(Array(1, 2, 3), 4, 5, "ABC") > 15 Then
    ///         ' won't run, error 13 "type mismatch" will be thrown when Variant/Error is compared to an Integer.
    ///     End If
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Private Sub Example()
    ///     Debug.Print Application.WorksheetFunction.Sum(Array(1, 2, 3), 4, 5, "ABC") ' raises error 1004
    /// 
    ///     Dim foo As Long
    ///     foo = Application.WorksheetFunction.Sum(Array(1, 2, 3), 4, 5, "ABC") ' raises error 1004
    /// 
    ///     If Application.WorksheetFunction.Sum(Array(1, 2, 3), 4, 5, "ABC") > 15 Then ' raises error 1004
    ///         ' won't run, error 1004 is raised when "ABC" is processed by WorksheetFunction.Sum, before it returns.
    ///     End If
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    [RequiredLibrary("Excel")]
    internal class ApplicationWorksheetFunctionInspection : IdentifierReferenceInspectionFromDeclarationsBase
    {
        public ApplicationWorksheetFunctionInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {}

        protected override IEnumerable<Declaration> ObjectionableDeclarations(DeclarationFinder finder)
        {
            var excel = finder.Projects.SingleOrDefault(item => !item.IsUserDefined && item.IdentifierName == "Excel");
            if (excel == null)
            {
                return Enumerable.Empty<Declaration>();
            }

            if (!(finder.FindClassModule("WorksheetFunction", excel, true) is ModuleDeclaration worksheetFunctionsModule))
            {
                return Enumerable.Empty<Declaration>();
            }

            if (!(finder.FindClassModule("Application", excel, true) is ModuleDeclaration excelApplicationClass))
            {
                return Enumerable.Empty<Declaration>();
            }

            var worksheetFunctionNames = worksheetFunctionsModule.Members
                .Where(decl => decl.DeclarationType == DeclarationType.Function)
                .Select(decl => decl.IdentifierName)
                .ToHashSet();

            return excelApplicationClass.Members
                .Where(decl => worksheetFunctionNames.Contains(decl.IdentifierName));
        }

        protected override string ResultDescription(IdentifierReference reference)
        {
            return string.Format(InspectionResults.ApplicationWorksheetFunctionInspection, reference.IdentifierName);
        }
    }
}
