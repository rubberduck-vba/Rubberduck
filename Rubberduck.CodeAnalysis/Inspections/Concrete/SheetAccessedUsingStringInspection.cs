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

namespace Rubberduck.Inspections.Concrete
{
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
            "_Global", "_Application", "Global", "Application", "Workbook"
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
                .Where(reference => !IsIgnoringInspectionResultFor(reference, AnnotationName) &&
                                    IsAccessedWithStringLiteralParameter(reference))
                .Select(reference => new IdentifierReferenceInspectionResult(this,
                    InspectionResults.SheetAccessedUsingStringInspection, State, reference)));

            var issues = new List<IdentifierReferenceInspectionResult>();

            foreach (var reference in references)
            {
                using (var component = GetVBComponentMatchingSheetName(reference)) 
                {
                    if (component == null)
                    {
                        continue;
                    }
                    using (var properties = component.Properties)
                    {
                        reference.Properties.CodeName = (string)properties.Single(property => property.Name == "CodeName").Value;
                    }
                    issues.Add(reference);
                }
            }
            return issues;
        }

        private static bool IsAccessedWithStringLiteralParameter(IdentifierReference reference)
        {
            // Second case accounts for global modules
            var indexExprContext = reference.Context.Parent.Parent as VBAParser.IndexExprContext ??
                                   reference.Context.Parent as VBAParser.IndexExprContext;

            var literalExprContext = indexExprContext
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
            var project = State.Projects.First(p => p.ProjectId == reference.QualifiedName.ProjectId);

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
