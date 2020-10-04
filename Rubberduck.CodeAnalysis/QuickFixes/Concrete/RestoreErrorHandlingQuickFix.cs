using System.Collections.Generic;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.InternalApi.Extensions;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Introduces an error-handling subroutine to ensure error state is properly handled on scope exit.
    /// </summary>
    /// <inspections>
    /// <inspection name="UnhandledOnErrorResumeNextInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="true" module="true" project="true" all="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     On Error Resume Next
    ///     Debug.Print ActiveWorkbook.FullName
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     On Error GoTo ErrHandler
    ///     Debug.Print ActiveWorkbook.FullName
    ///     Exit Sub
    /// ErrHandler:
    ///     If Err.Number > 0 Then 'TODO: handle specific error
    ///         Err.Clear
    ///         Resume Next
    ///     End If
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class RestoreErrorHandlingQuickFix : QuickFixBase
    {
        private const string LabelPrefix = "ErrorHandler";

        public RestoreErrorHandlingQuickFix()
            : base(typeof(UnhandledOnErrorResumeNextInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            if (!(result is IWithInspectionResultProperties<IReadOnlyList<VBAParser.OnErrorStmtContext>> resultProperties))
            {
                return;
            }

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
            var labels = bodyElementContext.GetDescendents<VBAParser.IdentifierStatementLabelContext>()
                .OrderBy(labelContext => labelContext.GetSelection())
                .ToArray();
            var maximumExistingLabelIndex = GetMaximumExistingLabelIndex(labels);
            var unhandledContexts = resultProperties.Properties;
            var offset = unhandledContexts.IndexOf(result.Context);
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

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;

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
