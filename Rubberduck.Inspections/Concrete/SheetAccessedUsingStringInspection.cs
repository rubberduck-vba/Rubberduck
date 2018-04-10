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
using Rubberduck.VBEditor;
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

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var worksheetsDeclarations = BuiltInDeclarations.Where(decl =>
                    decl.ProjectName.Equals("Excel") &&
                    decl.ComponentName.Equals("Workbook") &&
                    decl.ParentDeclaration.IdentifierName.Equals("Workbook") &&
                    (decl.IdentifierName.Equals("Worksheets") || decl.IdentifierName.Equals("Sheets")))
                .ToArray();

            var references = from declaration in worksheetsDeclarations
                    from reference in declaration.References.Where(r =>
                        !IsIgnoringInspectionResultFor(r, AnnotationName) && IsAccessedUsingThisWorkbook(r) && IsAccessedWithStringLiteralParameter(r))
                    let qualifiedSelection = new QualifiedSelection(reference.QualifiedModuleName, reference.Selection)
                    select new IdentifierReferenceInspectionResult(this,
                        InspectionsUI.SheetAccessedUsingStringInspectionResultFormat,
                        State,
                        reference);

            var issues = new List<IdentifierReferenceInspectionResult>();

            foreach (var reference in references)
            {
                var component = GetVBComponentMatchingSheetName(reference);
                if (component != null)
                {
                    reference.Properties.CodeName = (string)component.Properties.First(property => property.Name == "CodeName").Value;
                    issues.Add(reference);
                }
            }

            return issues;
        }

        private static bool IsAccessedUsingThisWorkbook(IdentifierReference reference)
        {
            return (reference.Context.Parent.GetChild(0) as VBAParser.SimpleNameExprContext)?.identifier().GetText() == "ThisWorkbook";
        }

        private static bool IsAccessedWithStringLiteralParameter(IdentifierReference reference)
        {
            return ((reference.Context.Parent.Parent as VBAParser.IndexExprContext)
                       ?.argumentList()
                       ?.argument(0)
                       ?.positionalArgument()
                       ?.argumentExpression().expression() as VBAParser.LiteralExprContext)
                   ?.literalExpression().STRINGLITERAL() != null;
        }

        private IVBComponent GetVBComponentMatchingSheetName(IdentifierReferenceInspectionResult reference)
        {
            var sheetArgumentContext = ((VBAParser.IndexExprContext)reference.Context.Parent.Parent).argumentList().argument(0);
            // Remove leading and trailing quotes
            var sheetName = string.Concat(sheetArgumentContext.GetText().Skip(1).Take(sheetArgumentContext.GetText().Length - 2));
            var project = State.Projects.First(p => p.ProjectId == reference.QualifiedName.ProjectId);

            return project.VBComponents.FirstOrDefault(c =>
                c.Type == ComponentType.Document &&
                (string) c.Properties.First(property => property.Name == "Name").Value == sheetName);
        }
    }
}
