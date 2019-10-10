using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.Inspections.Inspections.Extensions;

namespace Rubberduck.Inspections.Concrete
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
    /// <example hasResults="true">
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
    /// </example>
    /// <example hasResults="false">
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
    /// </example>
    [RequiredLibrary("Excel")]
    public class ApplicationWorksheetFunctionInspection : InspectionBase
    {
        public ApplicationWorksheetFunctionInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var excel = State.DeclarationFinder.Projects.SingleOrDefault(item => !item.IsUserDefined && item.IdentifierName == "Excel");
            if (excel == null) { return Enumerable.Empty<IInspectionResult>(); }

            var members = new HashSet<string>(BuiltInDeclarations.Where(decl => decl.DeclarationType == DeclarationType.Function &&
                                                                        decl.ParentDeclaration != null && 
                                                                        decl.ParentDeclaration.ComponentName.Equals("WorksheetFunction"))
                                                                 .Select(decl => decl.IdentifierName));

            var usages = BuiltInDeclarations.Where(decl => decl.References.Any() &&
                                                           decl.ProjectName.Equals("Excel") &&
                                                           decl.ComponentName.Equals("Application") &&
                                                           members.Contains(decl.IdentifierName));

            return from usage in usages
                   // filtering on references isn't the default ignore filtering
                   from reference in usage.References.Where(use => !use.IsIgnoringInspectionResultFor(AnnotationName))
                   let qualifiedSelection = new QualifiedSelection(reference.QualifiedModuleName, reference.Selection)
                   select new IdentifierReferenceInspectionResult(this,
                                               string.Format(InspectionResults.ApplicationWorksheetFunctionInspection, usage.IdentifierName),
                                               State,
                                               reference);
        }
    }
}
