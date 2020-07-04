using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.CodeAnalysis.QuickFixes.Concrete
{
    /// <summary>
    /// Removes an instruction that references a variable that isn't assigned. This operation may break the code.
    /// </summary>
    /// <inspections>
    /// <inspection name="UnassignedVariableUsageInspection" />
    /// </inspections>
    /// <canfix multiple="true" procedure="true" module="true" project="true" all="true" />
    /// <example>
    /// <before>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     Dim value As Long
    ///     Debug.Print value
    /// End Sub
    /// ]]>
    /// </before>
    /// <after>
    /// <![CDATA[
    /// Option Explicit
    /// 
    /// Public Sub DoSomething()
    ///     Dim value As Long
    ///     
    /// End Sub
    /// ]]>
    /// </after>
    /// </example>
    internal sealed class RemoveUnassignedVariableUsageQuickFix : QuickFixBase
    {
        public RemoveUnassignedVariableUsageQuickFix()
            : base(typeof(UnassignedVariableUsageInspection))
        {}

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName);

            if (result.Context.Parent.Parent is VBAParser.WithStmtContext withContext)
            {
                var lines = withContext.GetText().Replace("\r", string.Empty).Split('\n');
                // Assume that the End With is at the appropriate indentation level for the block. Note that this could
                // over-indent or under-indent some lines if statement separators are being used, but meh. 
                var padding = new string(' ', lines.Last().IndexOf(Tokens.End, StringComparison.Ordinal));

                var replacement = new List<string>
                {
                    $"{Tokens.CommentMarker}TODO - {result.Description}",
                    $"{Tokens.CommentMarker}{padding}{lines.First()}"
                };
                replacement.AddRange(lines.Skip(1)
                    .Select(line => Tokens.CommentMarker + line));
                
                rewriter.Replace(withContext, string.Join(Environment.NewLine, replacement));
                return;
            }
            var assignmentContext = result.Context.GetAncestor<VBAParser.LetStmtContext>() ??
                                    (ParserRuleContext)result.Context.GetAncestor<VBAParser.CallStmtContext>();

            rewriter.Remove(assignmentContext);
        }

        public override string Description(IInspectionResult result) => Resources.Inspections.QuickFixes.RemoveUnassignedVariableUsageQuickFix;

        public override bool CanFixMultiple => true;
        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
        public override bool CanFixAll => true;
    }
}