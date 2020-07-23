using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.CodeAnalysis.Inspections.Attributes;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Locates public User-Defined Function procedures accidentally named after a cell reference.
    /// </summary>
    /// <reference name="Excel" />
    /// <why>
    /// Another good reason to avoid numeric suffixes: if the function is meant to be used as a UDF in a cell formula,
    /// the worksheet cell by the same name takes precedence and gets the reference, and the function is never invoked.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Function FOO1234()
    /// End Function
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Function Foo()
    /// End Function
    /// ]]>
    /// </module>
    /// </example>
    [RequiredLibrary("Excel")]
    internal class ExcelUdfNameIsValidCellReferenceInspection : DeclarationInspectionUsingGlobalInformationBase<bool>
    {
        public ExcelUdfNameIsValidCellReferenceInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider, new []{DeclarationType.Function}, new []{DeclarationType.PropertyGet, DeclarationType.LibraryFunction})
        {}

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module, DeclarationFinder finder, bool excelIsReferenced)
        {
            if (!excelIsReferenced || module.ComponentType != ComponentType.StandardModule)
            {
                return Enumerable.Empty<IInspectionResult>();
            }

            var proceduralModuleDeclaration = finder.Members(module, DeclarationType.ProceduralModule)
                        .SingleOrDefault() as ProceduralModuleDeclaration;

            if (proceduralModuleDeclaration == null
                || proceduralModuleDeclaration.IsPrivateModule)
            {
                return Enumerable.Empty<IInspectionResult>();
            }

            return base.DoGetInspectionResults(module, finder, excelIsReferenced);
        }

        protected override bool GlobalInformation(DeclarationFinder finder)
        {
            return finder.Projects.Any(project => !project.IsUserDefined
                                                            && project.IdentifierName == "Excel");
        }

        protected override bool IsResultDeclaration(Declaration declaration, DeclarationFinder finder, bool globalInfo)
        {
            if (!VisibleAsUdf.Contains(declaration.Accessibility))
            {
                return false;
            }

            var cellIdMatch = ValidCellIdRegex.Match(declaration.IdentifierName);
            if (!cellIdMatch.Success)
            {
                return false;
            }

            var row = Convert.ToUInt32(cellIdMatch.Groups["Row"].Value);
            return row > 0 && row <= MaximumExcelRows;
        }

        private static readonly Regex ValidCellIdRegex =
            new Regex(@"^([a-z]|[a-z]{2}|[a-w][a-z]{2}|x([a-e][a-z]|f[a-d]))(?<Row>\d+)$",
                RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.ExplicitCapture);

        private static readonly HashSet<Accessibility> VisibleAsUdf = new HashSet<Accessibility> { Accessibility.Public, Accessibility.Implicit };

        private const uint MaximumExcelRows = 1048576;

        protected override string ResultDescription(Declaration declaration)
        {
            return string.Format(InspectionResults.ExcelUdfNameIsValidCellReferenceInspection, declaration.IdentifierName);
        }
    }
}
