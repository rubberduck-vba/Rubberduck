using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public class AccessSheetUsingCodeNameQuickFix : QuickFixBase
    {
        private readonly RubberduckParserState _state;

        public AccessSheetUsingCodeNameQuickFix(RubberduckParserState state)
            : base(typeof(SheetAccessedUsingStringInspection))
        {
            _state = state;
        }

        public override void Fix(IInspectionResult result)
        {
            var referenceResult = (IdentifierReferenceInspectionResult)result;
            var sheet = ((VBAParser.IndexExprContext)referenceResult.Context.Parent.Parent).argumentList().argument(0);
            // Get rid of leading and trailing quotes
            var sheetName = string.Concat(sheet.GetText().Skip(1).Take(sheet.GetText().Length - 2));

            var rewriter = _state.GetRewriter(referenceResult.QualifiedName);

            var setStatement = referenceResult.Context.GetAncestor<VBAParser.SetStmtContext>();
            if (setStatement == null)
            {
                // Sheet accessed inline

                rewriter.Replace(referenceResult.Context.Parent.Parent as ParserRuleContext, sheetName);
            }
            else
            {
                // Sheet assigned to variable

                var sheetVariableName = setStatement.lExpression().GetText();
                var sheetDeclaration = _state.DeclarationFinder.MatchName(sheetVariableName)
                    .First(declaration =>
                    {
                        var moduleBodyElement = declaration.Context.GetAncestor<VBAParser.ModuleBodyElementContext>();
                        return moduleBodyElement != null && moduleBodyElement == referenceResult.Context.GetAncestor<VBAParser.ModuleBodyElementContext>();
                    });

                var variableListContext = (VBAParser.VariableListStmtContext)sheetDeclaration.Context.Parent;
                if (variableListContext.variableSubStmt().Length == 1)
                {
                    rewriter.Remove(variableListContext.Parent as ParserRuleContext);
                }
                else if (sheetDeclaration.Context == variableListContext.variableSubStmt().Last())
                {
                    rewriter.Remove(variableListContext.COMMA().Last());
                    rewriter.Remove(sheetDeclaration);
                }
                else
                {
                    rewriter.Remove(variableListContext.COMMA().First(comma => comma.Symbol.StartIndex > sheetDeclaration.Context.Start.StartIndex));
                    rewriter.Remove(sheetDeclaration);
                }

                foreach (var reference in sheetDeclaration.References)
                {
                    rewriter.Replace(reference.Context, sheetName);
                }

                rewriter.Remove(setStatement);
            }
        }

        public override string Description(IInspectionResult result)
        {
            return InspectionsUI.AccessSheetUsingCodeNameQuickFix;
        }

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}
