using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class RemoveUnassignedVariableUsageQuickFix : QuickFixBase
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

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}