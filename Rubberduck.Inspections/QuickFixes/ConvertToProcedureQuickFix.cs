using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class ConvertToProcedureQuickFix : IQuickFix
    {
        private readonly RubberduckParserState _state;
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type>
        {
            typeof(NonReturningFunctionInspection),
            typeof(FunctionReturnValueNotUsedInspection)
        };

        public ConvertToProcedureQuickFix(RubberduckParserState state)
        {
            _state = state;
        }

        public IReadOnlyCollection<Type> SupportedInspections => _supportedInspections.ToList();

        public void Fix(IInspectionResult result)
        {
            var functionContext = result.Context as VBAParser.FunctionStmtContext;
            if (functionContext != null)
            {
                ConvertFunction(result, functionContext);
            }

            var propertyGetContext = result.Context as VBAParser.PropertyGetStmtContext;
            if (propertyGetContext != null)
            {
                ConvertPropertyGet(result, propertyGetContext);
            }
        }

        private void ConvertFunction(IInspectionResult result, VBAParser.FunctionStmtContext functionContext)
        {
            var rewriter = _state.GetRewriter(result.Target);

            var asTypeContext = ParserRuleContextHelper.GetChild<VBAParser.AsTypeClauseContext>(functionContext);
            if (asTypeContext != null)
            {
                rewriter.Remove(asTypeContext);
                rewriter.Remove(functionContext.children.ElementAt(functionContext.children.IndexOf(asTypeContext) - 1) as ParserRuleContext);
            }

            if (result.Target.TypeHint != null)
            {
                rewriter.Remove(ParserRuleContextHelper.GetDescendent<VBAParser.TypeHintContext>(functionContext));
            }

            rewriter.Replace(functionContext.FUNCTION(), Tokens.Sub);
            rewriter.Replace(functionContext.END_FUNCTION(), "End Sub");

            foreach (var returnStatement in GetReturnStatements(result.Target))
            {
                rewriter.Remove(returnStatement);
            }
        }

        private void ConvertPropertyGet(IInspectionResult result, VBAParser.PropertyGetStmtContext propertyGetContext)
        {
            var rewriter = _state.GetRewriter(result.Target);

            var asTypeContext = ParserRuleContextHelper.GetChild<VBAParser.AsTypeClauseContext>(propertyGetContext);
            if (asTypeContext != null)
            {
                rewriter.Remove(asTypeContext);
                rewriter.Remove(propertyGetContext.children.ElementAt(propertyGetContext.children.IndexOf(asTypeContext) - 1) as ParserRuleContext);
            }

            if (result.Target.TypeHint != null)
            {
                rewriter.Remove(ParserRuleContextHelper.GetDescendent<VBAParser.TypeHintContext>(propertyGetContext));
            }

            rewriter.Replace(propertyGetContext.PROPERTY_GET(), Tokens.Sub);
            rewriter.Replace(propertyGetContext.END_PROPERTY(), "End Sub");

            foreach (var returnStatement in GetReturnStatements(result.Target))
            {
                rewriter.Remove(returnStatement);
            }
        }

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.ConvertFunctionToProcedureQuickFix;
        }

        public bool CanFixInProcedure => false;
        public bool CanFixInModule => true;
        public bool CanFixInProject => false;

        private IEnumerable<ParserRuleContext> GetReturnStatements(Declaration declaration)
        {
            return declaration.References
                .Where(usage => IsReturnStatement(declaration, usage))
                .Select(usage => usage.Context.Parent)
                .Cast<ParserRuleContext>();
        }

        private bool IsReturnStatement(Declaration declaration, IdentifierReference assignment)
        {
            return assignment.ParentScoping.Equals(declaration) && assignment.Declaration.Equals(declaration);
        }
    }
}
