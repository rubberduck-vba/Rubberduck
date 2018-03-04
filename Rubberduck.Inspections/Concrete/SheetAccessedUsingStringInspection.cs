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

namespace Rubberduck.Inspections.Concrete
{
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

            return from declaration in worksheetsDeclarations
                from reference in declaration.References.Where(r =>
                    !IsIgnoringInspectionResultFor(r, AnnotationName) && IsAccessedUsingThisWorkbook(r) && IsAccessedWithStringLiteralParameter(r))
                let qualifiedSelection = new QualifiedSelection(reference.QualifiedModuleName, reference.Selection)
                select new IdentifierReferenceInspectionResult(this,
                    InspectionsUI.SheetAccessedUsingStringInspectionResultFormat,
                    State,
                    reference);
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
    }
}
