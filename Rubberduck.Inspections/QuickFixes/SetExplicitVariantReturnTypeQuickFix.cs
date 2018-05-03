using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class SetExplicitVariantReturnTypeQuickFix : QuickFixBase
    {
        private readonly RubberduckParserState _state;

        public SetExplicitVariantReturnTypeQuickFix(RubberduckParserState state)
            :base(typeof(ImplicitVariantReturnTypeInspection))
        {
            _state = state;
        }

        public override void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.Target);
            
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

        public override string Description(IInspectionResult result) => InspectionsUI.SetExplicitVariantReturnTypeQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}