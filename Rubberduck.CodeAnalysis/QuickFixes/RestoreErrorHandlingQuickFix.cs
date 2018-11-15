using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class RestoreErrorHandlingQuickFix : QuickFixBase
    {
        private const string LabelPrefix = "ErrorHandler";

        public RestoreErrorHandlingQuickFix()
            : base(typeof(UnhandledOnErrorResumeNextInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var exitStatement = "Exit ";
            VBAParser.BlockContext block;
            var bodyElementContext = result.Context.GetAncestor<VBAParser.ModuleBodyElementContext>();

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

            var rewriter = rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName);
            var context = (VBAParser.OnErrorStmtContext)result.Context;
            var labels = bodyElementContext.GetDescendents<VBAParser.IdentifierStatementLabelContext>().ToArray();
            var maximumExistingLabelIndex = GetMaximumExistingLabelIndex(labels);
            int offset = result.Properties.UnhandledContexts.IndexOf(result.Context);
            var labelIndex = maximumExistingLabelIndex + offset;

            var labelSuffix = labelIndex == 0
                ? labels.Select(GetLabelText).Any(text => text == LabelPrefix)
                    ? "1"
                    : ""
                : maximumExistingLabelIndex == 0
                    ? labelIndex.ToString()
                    : (labelIndex + 1).ToString();

            rewriter.Replace(context.RESUME(), Tokens.GoTo);
            rewriter.Replace(context.NEXT(), $"{LabelPrefix}{labelSuffix}");

            var errorHandlerSubroutine = $@"
    {exitStatement}
{LabelPrefix}{labelSuffix}:
    If Err.Number > 0 Then 'TODO: handle specific error
        Err.Clear
        Resume Next
    End If
";

            rewriter.InsertAfter(block.Stop.TokenIndex, errorHandlerSubroutine);
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.UnhandledOnErrorResumeNextInspectionQuickFix;

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;

        private static int GetMaximumExistingLabelIndex(IEnumerable<VBAParser.IdentifierStatementLabelContext> labelContexts)
        {
            var maximumIndex = 0;

            foreach (var context in labelContexts)
            {
                var labelText = GetLabelText(context);
                if (labelText.ToLower().StartsWith(LabelPrefix.ToLower()))
                {
                    var suffixIsNumeric = int.TryParse(string.Concat(labelText.Skip(LabelPrefix.Length)), out var index);
                    if (suffixIsNumeric && index > maximumIndex)
                    {
                        maximumIndex = index;
                    }
                }
            }

            return maximumIndex;
        }

        private static string GetLabelText(VBAParser.IdentifierStatementLabelContext labelContext)
        {
            return labelContext.legalLabelIdentifier().identifier().untypedIdentifier().identifierValue().IDENTIFIER().GetText();
        }
    }
}
