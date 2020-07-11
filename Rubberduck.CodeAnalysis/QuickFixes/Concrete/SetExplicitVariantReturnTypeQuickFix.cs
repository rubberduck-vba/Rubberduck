using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Declares an explicit Variant return type for a procedure that implicitly returns a Variant.
    /// </summary>
    /// <inspections>
    /// <inspection name="ImplicitVariantReturnTypeInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="true" module="true" project="true" all="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Function WorksheetErrorValueNA()
    ///     WorksheetErrorValueNA = CVErr(xlErrNA)
    /// End Function
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Function WorksheetErrorValueNA() As Variant
    ///     WorksheetErrorValueNA = CVErr(xlErrNA)
    /// End Function
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class SetExplicitVariantReturnTypeQuickFix : QuickFixBase
    {
        public SetExplicitVariantReturnTypeQuickFix()
            :base(typeof(ImplicitVariantReturnTypeInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.Target.QualifiedModuleName);
            
            const string asTypeClause = " As Variant";
            switch (result.Target.DeclarationType)
            {
                case DeclarationType.Variable:
                    var variableContext = (VBAParser.VariableSubStmtContext)result.Target.Context;
                    rewriter.InsertAfter(variableContext.identifier().Stop.TokenIndex, asTypeClause);
                    break;
                case DeclarationType.Parameter:
                    var parameterContext = (VBAParser.ArgContext)result.Target.Context;
                    rewriter.InsertAfter(parameterContext.unrestrictedIdentifier().Stop.TokenIndex, asTypeClause);
                    break;
                case DeclarationType.Function:
                    var functionContext = (VBAParser.FunctionStmtContext)result.Target.Context;
                    rewriter.InsertAfter(functionContext.argList().Stop.TokenIndex, asTypeClause);
                    break;
                case DeclarationType.LibraryFunction:
                    var declareContext = (VBAParser.DeclareStmtContext)result.Target.Context;
                    rewriter.InsertAfter(declareContext.argList().Stop.TokenIndex, asTypeClause);
                    break;
                case DeclarationType.PropertyGet:
                    var propertyContext = (VBAParser.PropertyGetStmtContext)result.Target.Context;
                    rewriter.InsertAfter(propertyContext.argList().Stop.TokenIndex, asTypeClause);
                    break;
            }
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.SetExplicitVariantReturnTypeQuickFix;

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;
    }
}