using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class AccessSheetUsingCodeNameQuickFix : QuickFixBase
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public AccessSheetUsingCodeNameQuickFix(IDeclarationFinderProvider declarationFinderProvider)
            : base(typeof(SheetAccessedUsingStringInspection))
        {
            _declarationFinderProvider = declarationFinderProvider;
        }

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var referenceResult = (IdentifierReferenceInspectionResult)result;

            var rewriter = rewriteSession.CheckOutModuleRewriter(referenceResult.QualifiedName);

            var setStatement = referenceResult.Context.GetAncestor<VBAParser.SetStmtContext>();
            var isArgument = referenceResult.Context.GetAncestor<VBAParser.ArgumentContext>() != null;
            if (setStatement == null || isArgument)
            {
                // Sheet accessed inline

                // Second case accounts for global modules
                var indexExprContext = referenceResult.Context.Parent.Parent as VBAParser.IndexExprContext ??
                                       referenceResult.Context.Parent as VBAParser.IndexExprContext;

                rewriter.Replace(indexExprContext, (string)referenceResult.Properties.CodeName);
            }
            else
            {
                // Sheet assigned to variable

                var sheetVariableName = setStatement.lExpression().GetText();
                var sheetDeclaration = _declarationFinderProvider.DeclarationFinder.MatchName(sheetVariableName)
                    .First(declaration =>
                    {
                        var moduleBodyElement = declaration.Context.GetAncestor<VBAParser.ModuleBodyElementContext>();
                        return moduleBodyElement != null && moduleBodyElement == referenceResult.Context.GetAncestor<VBAParser.ModuleBodyElementContext>();
                    });

                if (!sheetDeclaration.IsUndeclared)
                {
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
                }

                foreach (var reference in sheetDeclaration.References)
                {
                    rewriter.Replace(reference.Context, (string)referenceResult.Properties.CodeName);
                }

                rewriter.Remove(setStatement);
            }
        }

        public override string Description(IInspectionResult result)
        {
            return Resources.Inspections.QuickFixes.AccessSheetUsingCodeNameQuickFix;
        }

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}
