using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.Inspections.Inspections.Extensions;
using Rubberduck.Parsing;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Locates ThisWorkbook.Worksheets and ThisWorkbook.Sheets calls that appear to be dereferencing a worksheet that is already accessible at compile-time with a global-scope identifier.
    /// </summary>
    /// <why>
    /// Sheet names can be changed by the user, as can a worksheet's index in ThisWorkbook.Worksheets. 
    /// Worksheets that exist in ThisWorkbook at compile-time are more reliably programmatically accessed using their CodeName, 
    /// which cannot be altered by the user without accessing the VBE and altering the VBA project.
    /// </why>
    /// <reference name="Excel" />
    /// <remarks>
    /// Inspection only evaluates hard-coded string literals; string-valued expressions evaluating into a sheet name are ignored.
    /// </remarks>
    /// <example>
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim sheet As Worksheet
    ///     Set sheet = ThisWorkbook.Worksheets("Sheet1") ' Sheet "Sheet1" exists at compile-time
    ///     sheet.Range("A1").Value = 42
    /// End Sub
    /// ]]>
    /// </example>
    /// <example>
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Sheet1.Range("A1").Value = 42 ' TODO rename Sheet1 to meaningful name
    /// End Sub
    /// ]]>
    /// </example>
    [RequiredHost("EXCEL.EXE")]
    [RequiredLibrary("Excel")]
    public class SheetAccessedUsingStringInspection : InspectionBase
    {
        public SheetAccessedUsingStringInspection(RubberduckParserState state) : base(state) { }

        private static readonly string[] InterestingMembers =
        {
            "Worksheets", "Sheets"
        };

        private static readonly string[] InterestingClasses =
        {
            "Workbook", "ThisWorkbook" // "_Global", "_Application", "Global", "Application", "Workbook"
        };

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var excel = State.DeclarationFinder.Projects.SingleOrDefault(item => !item.IsUserDefined && item.IdentifierName == "Excel");
            if (excel == null)
            {
                return Enumerable.Empty<IInspectionResult>();
            }

            var targetProperties = BuiltInDeclarations
                .OfType<PropertyDeclaration>()
                .Where(x => InterestingMembers.Contains(x.IdentifierName) && InterestingClasses.Contains(x.ParentDeclaration?.IdentifierName))
                .ToList();

            var references = targetProperties.SelectMany(declaration => declaration.References
                .Where(reference => !reference.IsIgnoringInspectionResultFor(AnnotationName) &&
                                    IsAccessedWithStringLiteralParameter(reference))).ToArray();

            var issues = references.Select(reference => 
                    new IdentifierReferenceInspectionResult(
                        this, InspectionResults.SheetAccessedUsingStringInspection, State, reference))
                .ToList();

            foreach (var issue in issues)
            {
                using (var component = GetVBComponentMatchingSheetName(issue)) 
                {
                    if (component == null)
                    {
                        continue;
                    }
                    using (var properties = component.Properties)
                    {
                        issue.Properties.CodeName = (string)properties.Single(property => property.Name == "CodeName").Value;
                    }
                    issues.Add(issue);
                }
            }
            return issues;
        }

        private static bool IsAccessedWithStringLiteralParameter(IdentifierReference reference)
        {
            if (reference.Context.Parent.Parent is VBAParser.IndexExprContext qualifiedMemberCall)
            {
                return HasStringLiteralParameter(qualifiedMemberCall)
                       && qualifiedMemberCall
                           ?.GetAncestor<VBAParser.MemberAccessExprContext>()
                           ?.lExpression().GetChild<VBAParser.SimpleNameExprContext>()
                           ?.GetText() == "ThisWorkbook";
            }

            return reference.Context.Parent is VBAParser.IndexExprContext unqualifiedMemberCall
                   && HasStringLiteralParameter(unqualifiedMemberCall)
                   && reference.ParentScoping.QualifiedModuleName.ComponentType == ComponentType.Document
                   && reference.ParentScoping.QualifiedModuleName.ComponentName == "ThisWorkbook";
        }

        private static bool HasStringLiteralParameter(VBAParser.IndexExprContext context)
        {
            var literalExprContext = context
                ?.argumentList()
                ?.argument(0)
                ?.positionalArgument()
                ?.argumentExpression().expression() as VBAParser.LiteralExprContext;

            return literalExprContext?.literalExpression().STRINGLITERAL() != null;
        }

        private IVBComponent GetVBComponentMatchingSheetName(IdentifierReferenceInspectionResult reference)
        {
            // Second case accounts for global modules
            var indexExprContext = reference.Context.Parent.Parent as VBAParser.IndexExprContext ??
                                   reference.Context.Parent as VBAParser.IndexExprContext;

            if (indexExprContext == null)
            {
                return null;
            }

            var sheetArgumentContext = indexExprContext.argumentList().argument(0);
            var sheetName = FormatSheetName(sheetArgumentContext.GetText());
            var project = State.ProjectsProvider.Project(reference.QualifiedName.ProjectId);

            using (var components = project.VBComponents)
            {
                foreach (var component in components)
                {
                    using (var properties = component.Properties)
                    {
                        if (component.Type != ComponentType.Document)
                        {
                            component.Dispose();
                            continue;
                        }
                        foreach (var property in properties)
                        {
                            var found = property.Name.Equals("Name") && ((string)property.Value).Equals(sheetName);
                            property.Dispose();
                            if (found)
                            {
                                return component;
                            }                          
                        }
                    }
                    component.Dispose();
                }
                return null;
            }
        }

        private static string FormatSheetName(string sheetName)
        {
            return sheetName.StartsWith("\"") && sheetName.EndsWith("\"")
                ? sheetName.Substring(1, sheetName.Length - 2)
                : sheetName;
        }
    }
}
