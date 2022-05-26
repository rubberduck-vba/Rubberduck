using System.Collections.Generic;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.CodeAnalysis.Inspections.Attributes;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
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
    /// <hostApp name="EXCEL.EXE" />
    /// <remarks>
    /// For performance reasons, the inspection only evaluates hard-coded string literals; string-valued expressions evaluating into a sheet name are ignored.
    /// </remarks>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim sheet As Worksheet
    ///     Set sheet = ThisWorkbook.Worksheets("Sheet1") ' Sheet "Sheet1" exists at compile-time
    ///     sheet.Range("A1").Value = 42
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Sheet1.Range("A1").Value = 42 ' TODO rename Sheet1 to meaningful name
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    [RequiredHost("EXCEL.EXE")]
    [RequiredLibrary("Excel")]
    internal class SheetAccessedUsingStringInspection : IdentifierReferenceInspectionFromDeclarationsBase<string>
    {
        private readonly IProjectsProvider _projectsProvider;

        public SheetAccessedUsingStringInspection(IDeclarationFinderProvider declarationFinderProvider, IProjectsProvider projectsProvider)
            : base(declarationFinderProvider)
        {
            _projectsProvider = projectsProvider;
        }

        /// <summary>
        /// We're interested in both explicitly and implicitly bound retrievals from a Sheets collection.
        /// </summary>
        private static readonly string[] InterestingProperties =
        {
            "Item", // explicit member call
            "_Default", // default member call (usually implicit)
        };

        protected override IEnumerable<Declaration> ObjectionableDeclarations(DeclarationFinder finder)
        {
            if (!finder.TryFindProjectDeclaration("Excel", out var excel))
            {
                // [RequiredHost] attribute puts this in "should not happen" territory.
                yield break;
            }
            var sheetsClass = (ModuleDeclaration)finder.FindClassModule("Sheets", excel, true);
            if (sheetsClass == null)
            {
                // [RequiredHost] attribute puts this in "should not happen" territory.
                yield break;
            }

            if (sheetsClass != null)
            {
                foreach (var property in sheetsClass.Members.OfType<PropertyDeclaration>())
                {
                    if (InterestingProperties.Any(name => name.Equals(property.IdentifierName, System.StringComparison.InvariantCultureIgnoreCase)))
                    {
                        yield return property;
                    }
                }
            }
        }

        private static ClassModuleDeclaration GetHostWorkbookDeclaration(DeclarationFinder finder)
        {
            var documentModuleQMNs = finder.AllModules.Where(m => m.ComponentType == ComponentType.Document);
            ClassModuleDeclaration result = null;
            foreach (var qmn in documentModuleQMNs)
            {
                var declaration = finder.ModuleDeclaration(qmn) as ClassModuleDeclaration;
                if (declaration.Supertypes.Any(t => t.IdentifierName.Equals("Workbook") && t.ProjectName == "Excel" && !t.IsUserDefined))
                {
                    result = declaration;
                    break;
                }
            }

            return result ?? throw new System.InvalidOperationException("Failed to find the host Workbook declaration.");
        }

        private static ClassModuleDeclaration GetHostApplicationDeclaration(DeclarationFinder finder)
        {
            var result = finder.MatchName("Application").OfType<ClassModuleDeclaration>().FirstOrDefault(t => t.ProjectName == "Excel" && !t.IsUserDefined) as ClassModuleDeclaration;
            return result ?? throw new System.InvalidOperationException("Failed to find the host Application declaration.");
        }

        protected override (bool isResult, string properties) IsResultReferenceWithAdditionalProperties(IdentifierReference reference, DeclarationFinder finder)
        {
            if (reference.IdentifierName.Equals(Tokens.Me, System.StringComparison.InvariantCultureIgnoreCase))
            {
                // if Me is a worksheet module, 
                return (false, null);
            }

            var hostWorkbookDeclaration = GetHostWorkbookDeclaration(finder);

            var context = reference.Context as VBAParser.MemberAccessExprContext 
                        ?? reference.Context.Parent as VBAParser.MemberAccessExprContext
                        ?? reference.Context.Parent.Parent as VBAParser.MemberAccessExprContext;

            if (context is VBAParser.MemberAccessExprContext memberAccess)
            {
                var appObjectDeclaration = GetHostApplicationDeclaration(finder);
                var isApplicationQualifier = appObjectDeclaration.References.Any(appRef => 
                    context.GetSelection().Contains(appRef.Selection) 
                    && appRef.QualifiedModuleName.Equals(reference.QualifiedModuleName));

                if (isApplicationQualifier)
                {
                    // Application.Sheets(...) is referring to the ActiveWorkbook, not necessarily ThisWorkbook.
                    return (false, null);
                }
            }

            var isHostWorkbookQualifier = hostWorkbookDeclaration.References.Any(thisWorkbookRef =>
                    context.GetSelection().Contains(thisWorkbookRef.Selection)
                    && thisWorkbookRef.QualifiedModuleName.Equals(reference.QualifiedModuleName));

            var parentModule = finder.ModuleDeclaration(reference.QualifiedModuleName);
            if (!isHostWorkbookQualifier && parentModule is ProceduralModuleDeclaration)
            {
                // in a standard module the reference is against ActiveWorkbook unless it's explicitly against ThisWorkbook.
                return (false, null);
            }

            var sheetNameArgumentLiteralExpressionContext = SheetNameArgumentLiteralExpressionContext(reference);
            if (sheetNameArgumentLiteralExpressionContext?.STRINGLITERAL() == null)
            {
                return (false, null);
            }

            var projectId = reference.QualifiedModuleName.ProjectId;
            var sheetName = sheetNameArgumentLiteralExpressionContext.GetText().FromVbaStringLiteral();
            var codeName = CodeNameOfVBComponentMatchingSheetName(projectId, sheetName);

            if (codeName == null)
            {
                return (false, null);
            }

            return (true, codeName);
        }

        private static VBAParser.LiteralExpressionContext SheetNameArgumentLiteralExpressionContext(IdentifierReference reference)
        {
            // Second case accounts for global modules
            var indexExprContext = reference.Context.Parent.Parent as VBAParser.IndexExprContext ??
                                   reference.Context.Parent as VBAParser.IndexExprContext;

            return (indexExprContext
                ?.argumentList()
                ?.argument(0)
                ?.positionalArgument()
                ?.argumentExpression()
                ?.expression() as VBAParser.LiteralExprContext)
                ?.literalExpression();
        }

        private string CodeNameOfVBComponentMatchingSheetName(string projectId, string sheetName)
        {
            var components = _projectsProvider.Components(projectId);
            foreach (var (_, component) in components)
            {
                if (component.Type != ComponentType.Document)
                {
                    continue;
                }

                var name = ComponentPropertyValue(component, "Name");
                if (sheetName.Equals(name))
                {
                    return ComponentPropertyValue(component, "CodeName");
                }
            }

            return null;
        }

        private static string ComponentPropertyValue(IVBComponent component, string propertyName)
        {
            using (var properties = component.Properties)
            {
                foreach (var property in properties)
                {
                    using (property)
                    {
                        if (property.Name == propertyName)
                        {
                            return (string)property.Value;
                        }
                    }
                }
            }

            return null;
        }

        protected override string ResultDescription(IdentifierReference reference, string codeName) => InspectionResults.SheetAccessedUsingStringInspection;
    }
}
