using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Resources.Inspections;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete.ThunderCode
{
    /// <summary hidden="true">
    /// A ThunderCode inspection that locates negative line numbers.
    /// </summary>
    /// <why>
    /// This inpection is flagging code we dubbed "ThunderCode", 
    /// code our friend Andrew Jackson would have written to confuse Rubberduck's parser and/or resolver. 
    /// The VBE does allow rather strange and unbelievable things to happen.
    /// </why>
    internal sealed class NegativeLineNumberInspection : ParseTreeInspectionBase<ParserRuleContext>
    {
        public NegativeLineNumberInspection(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {
            ContextListener = new NegativeLineNumberKeywordsListener();
        }

        protected override IInspectionListener<ParserRuleContext> ContextListener { get; }

        protected override string ResultDescription(QualifiedContext<ParserRuleContext> context)
        {
            return InspectionResults.NegativeLineNumberInspection.ThunderCodeFormat();
        }

        protected override bool IsResultContext(QualifiedContext<ParserRuleContext> context, DeclarationFinder finder)
        {
            return !IsOnErrorGotoMinusOne(context.Context) 
                || ProcedureHasMinusOneLabel(finder, context);
        }

        private static bool IsOnErrorGotoMinusOne(ParserRuleContext context)
        {
            return context is VBAParser.OnErrorStmtContext onErrorStatement
                   && "-1".Equals(onErrorStatement.expression()?.GetText().Trim());
        }

        private static bool ProcedureHasMinusOneLabel(DeclarationFinder finder, QualifiedContext<ParserRuleContext> context)
        {
            return finder.Members(context.ModuleName, DeclarationType.LineLabel)
                .Any(label => label.IdentifierName.Equals("-1")
                    && (label.ParentScopeDeclaration
                        .Context?.GetSelection()
                        .Contains(context.Context.GetSelection()) ?? false));
        }

        private class NegativeLineNumberKeywordsListener : InspectionListenerBase<ParserRuleContext>
        {
            public override void EnterOnErrorStmt(VBAParser.OnErrorStmtContext context)
            {
                CheckContext(context, context.expression());
                base.EnterOnErrorStmt(context);
            }

            public override void EnterGoToStmt(VBAParser.GoToStmtContext context)
            {
                CheckContext(context, context.expression());
                base.EnterGoToStmt(context);
            }

            public override void EnterLineNumberLabel(VBAParser.LineNumberLabelContext context)
            {
                CheckContext(context, context);
                base.EnterLineNumberLabel(context);
            }

            private void CheckContext(ParserRuleContext context, IParseTree expression)
            {
                var target = expression?.GetText().Trim() ?? string.Empty;
                if (target.StartsWith("-") && int.TryParse(target.Substring(1), out _))
                {
                    SaveContext(context);
                }
            }
        }
    }
}
