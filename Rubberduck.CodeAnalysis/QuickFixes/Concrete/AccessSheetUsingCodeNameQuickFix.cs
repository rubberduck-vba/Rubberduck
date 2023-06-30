using System.Linq;
using Antlr4.Runtime;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.Inspections.Results;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Modifies a Workbook.Worksheets or Workbook.Sheets call accessing a sheet of ThisWorkbook that exists at compile-time.
    /// </summary>
    /// <inspections>
    /// <inspection name="SheetAccessedUsingStringInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="true" module="true" project="true" all="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim sheet As Worksheet
    ///     Set sheet = ThisWorkbook.Sheets("Sheet1")
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim sheet As Worksheet
    ///     Set sheet = Sheet1 '<~ note: local variable becomes redundant
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class AccessSheetUsingCodeNameQuickFix : QuickFixBase
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public AccessSheetUsingCodeNameQuickFix(IDeclarationFinderProvider declarationFinderProvider)
            : base(typeof(SheetAccessedUsingStringInspection))
        {
            _declarationFinderProvider = declarationFinderProvider;
        }

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var referenceResult = (IdentifierReferenceInspectionResult<string>)result;

            var rewriter = rewriteSession.CheckOutModuleRewriter(referenceResult.QualifiedName);

            var setStatement = referenceResult.Context.GetAncestor<VBAParser.SetStmtContext>();
            var isArgument = referenceResult.Context.GetAncestor<VBAParser.ArgumentContext>() != null;
            if (setStatement == null || isArgument)
            {
                // Sheet accessed inline

                // Second case accounts for global modules
                var indexExprContext = referenceResult.Context.Parent.Parent as VBAParser.IndexExprContext ??
                                       referenceResult.Context.Parent as VBAParser.IndexExprContext;

                var codeName = referenceResult.Properties;
                rewriter.Replace(indexExprContext, codeName);
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
                    var codeName = referenceResult.Properties;
                    rewriter.Replace(reference.Context, codeName);
                }

                rewriter.Remove(setStatement);
            }
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.AccessSheetUsingCodeNameQuickFix;

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;
    }
}
