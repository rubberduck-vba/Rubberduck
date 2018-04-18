using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
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
        public SheetAccessedUsingStringInspection(RubberduckParserState state) : base(state)
        {
        }

        private static readonly string[] Targets =
        {
            "Worksheets", "Sheets"
        };

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var excel = State.DeclarationFinder.Projects.SingleOrDefault(item => !item.IsUserDefined && item.IdentifierName == "Excel");
            if (excel == null)
            {
                return Enumerable.Empty<IInspectionResult>();
                
            }

            var modules = new[]
            {
                State.DeclarationFinder.FindClassModule("_Global", excel, true),
                State.DeclarationFinder.FindClassModule("_Application", excel, true),
                State.DeclarationFinder.FindClassModule("Global", excel, true),
                State.DeclarationFinder.FindClassModule("Application", excel, true),
                State.DeclarationFinder.FindClassModule("Workbook", excel, true),
            };

            var references = Targets
                .SelectMany(target => modules.SelectMany(module => State.DeclarationFinder.FindMemberMatches(module, target)))
                .Where(declaration => declaration.References.Any())
                .SelectMany(declaration => declaration.References
                    .Where(reference =>
                        !IsIgnoringInspectionResultFor(reference, AnnotationName) && IsAccessedWithStringLiteralParameter(reference))
                    .Select(reference => new IdentifierReferenceInspectionResult(this,
                        InspectionsUI.SheetAccessedUsingStringInspectionResultFormat, State, reference)));

            var issues = new List<IdentifierReferenceInspectionResult>();

            foreach (var reference in references)
            {
                var component = GetVBComponentMatchingSheetName(reference);
                if (component != null)
                {
                    reference.Properties.CodeName = (string)component.Properties.Single(property => property.Name == "CodeName").Value;
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

            return project.VBComponents.FirstOrDefault(c =>
                c.Type == ComponentType.Document &&
                (string) c.Properties.First(property => property.Name == "Name").Value == sheetName);
        }

        private static string FormatSheetName(string sheetName)
        {
            var formattedName = sheetName.First() == '"' ? sheetName.Skip(1) : sheetName;

            if (sheetName.Last() == '"')
            {
                formattedName = formattedName.Take(formattedName.Count() - 1);
            }

            return string.Concat(formattedName);
        }
    }
}
