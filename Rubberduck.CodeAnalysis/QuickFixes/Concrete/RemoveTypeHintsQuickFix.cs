using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Removes type hint characters from identifier declarations and value literals.
    /// </summary>
    /// <inspections>
    /// <inspection name="ObsoleteTypeHintInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="true" module="true" project="true" all="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     Dim message$
    ///     message$ = "Hi"
    ///     MsgBox message$
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     Dim message As String
    ///     message = "Hi"
    ///     MsgBox message
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class RemoveTypeHintsQuickFix : QuickFixBase
    {
        public RemoveTypeHintsQuickFix()
            : base(typeof(ObsoleteTypeHintInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            if (!string.IsNullOrWhiteSpace(result.Target.TypeHint))
            {
                var rewriter = rewriteSession.CheckOutModuleRewriter(result.Target.QualifiedModuleName);
                var typeHintContext = result.Context.GetDescendent<VBAParser.TypeHintContext>();

                rewriter.Remove(typeHintContext);

                var asTypeClause = ' ' + Tokens.As + ' ' + SymbolList.TypeHintToTypeName[result.Target.TypeHint];
                switch (result.Target.DeclarationType)
                {
                    case DeclarationType.Variable:
                        var variableContext = (VBAParser.VariableSubStmtContext) result.Target.Context;
                        rewriter.InsertAfter(variableContext.identifier().Stop.TokenIndex, asTypeClause);
                        break;
                    case DeclarationType.Constant:
                        var constantContext = (VBAParser.ConstSubStmtContext) result.Target.Context;
                        rewriter.InsertAfter(constantContext.identifier().Stop.TokenIndex, asTypeClause);
                        break;
                    case DeclarationType.Parameter:
                        var parameterContext = (VBAParser.ArgContext)result.Target.Context;
                        rewriter.InsertAfter(parameterContext.unrestrictedIdentifier().Stop.TokenIndex, asTypeClause);
                        break;
                    case DeclarationType.Function:
                        var functionContext = (VBAParser.FunctionStmtContext) result.Target.Context;
                        rewriter.InsertAfter(functionContext.argList().Stop.TokenIndex, asTypeClause);
                        break;
                    case DeclarationType.PropertyGet:
                        var propertyContext = (VBAParser.PropertyGetStmtContext)result.Target.Context;
                        rewriter.InsertAfter(propertyContext.argList().Stop.TokenIndex, asTypeClause);
                        break;
                }
            }

            foreach (var reference in result.Target.References)
            {
                var context = reference.Context.GetDescendent<VBAParser.TypeHintContext>();

                if (context != null)
                {
                    var rewriter = rewriteSession.CheckOutModuleRewriter(reference.QualifiedModuleName);
                    rewriter.Remove(context);
                }
            }
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.RemoveTypeHintsQuickFix;

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;
    }
}