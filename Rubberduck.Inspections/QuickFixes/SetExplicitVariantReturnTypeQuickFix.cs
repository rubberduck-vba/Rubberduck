using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class SetExplicitVariantReturnTypeQuickFix : QuickFixBase, IQuickFix
    {
        private readonly RubberduckParserState _state;

        public SetExplicitVariantReturnTypeQuickFix(RubberduckParserState state, InspectionLocator inspectionLocator)
        {
            _state = state;
            RegisterInspections(inspectionLocator.GetInspection<ImplicitVariantReturnTypeInspection>());
        }

        public void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.Target);
            
            var asTypeClause = " As Variant";
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

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.SetExplicitVariantReturnTypeQuickFix;
        }

        public bool CanFixInProcedure => true;
        public bool CanFixInModule => true;
        public bool CanFixInProject => true;
    }
}