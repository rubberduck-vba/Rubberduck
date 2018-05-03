using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public class RestoreErrorHandlingQuickFix : QuickFixBase
    {
        private readonly RubberduckParserState _state;

        public RestoreErrorHandlingQuickFix(RubberduckParserState state)
            : base(typeof(UnhandledOnErrorResumeNextInspection))
        {
            _state = state;
        }

        public override void Fix(IInspectionResult result)
        {
            var rewriter = _state.GetRewriter(result.QualifiedSelection.QualifiedName);
            var context = (VBAParser.OnErrorStmtContext)result.Context;

            rewriter.Replace(context.RESUME(), Tokens.GoTo);
            rewriter.Replace(context.NEXT(), result.Properties.Label);

            var exitStatement = "Exit ";
            VBAParser.BlockContext block;
            VBAParser.ModuleBodyElementContext bodyElementContext = result.Properties.BodyElement;

            if (bodyElementContext.propertyGetStmt() != null)
            {
                exitStatement += "Property";
                block = bodyElementContext.propertyGetStmt().block();
            }
            else if (bodyElementContext.propertyLetStmt() != null)
            {
                exitStatement += "Property";
                block = bodyElementContext.propertyLetStmt().block();
            }
            else if (bodyElementContext.propertySetStmt() != null)
            {
                exitStatement += "Property";
                block = bodyElementContext.propertySetStmt().block();
            }
            else if (bodyElementContext.functionStmt() != null)
            {
                exitStatement += "Function";
                block = bodyElementContext.functionStmt().block();
            }
            else
            {
                exitStatement += "Sub";
                block = bodyElementContext.subStmt().block();
            }

            var errorHandlerSubroutine = $@"
    {exitStatement}
{result.Properties.Label}:
    If Err.Number > 0 Then 'TODO: handle specific error
        Err.Clear
        Resume Next
    End If
";

            rewriter.InsertAfter(block.Stop.TokenIndex, errorHandlerSubroutine);
        }

        public override string Description(IInspectionResult result) => InspectionsUI.UnhandledOnErrorResumeNextInspectionQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}
