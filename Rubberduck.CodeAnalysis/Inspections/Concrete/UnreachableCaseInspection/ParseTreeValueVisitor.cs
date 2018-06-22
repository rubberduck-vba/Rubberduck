using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IParseTreeValueVisitor : IParseTreeVisitor<IParseTreeVisitorResults>
    {
        event EventHandler<ValueResultEventArgs> OnValueResultCreated;
    }

    public class ParseTreeValueVisitor : IParseTreeValueVisitor
    {
        private IParseTreeVisitorResults _contextValues;
        private IParseTreeValueFactory _inspValueFactory;

        public ParseTreeValueVisitor(IParseTreeValueFactory valueFactory, Func<ParserRuleContext, (bool success, IdentifierReference idRef)> idRefRetriever)
        {
            _inspValueFactory = valueFactory;
            IdRefRetriever = idRefRetriever;
            _contextValues = new ParseTreeVisitorResults();
            OnValueResultCreated += _contextValues.OnNewValueResult;
        }

        //used only by UnreachableCaseInspection tests
        public RubberduckParserState State { set; get; } = null;

        private Func<ParserRuleContext, (bool success, IdentifierReference idRef)> IdRefRetriever { set; get; } = null;

        public event EventHandler<ValueResultEventArgs> OnValueResultCreated;

        public virtual IParseTreeVisitorResults Visit(IParseTree tree)
        {
            if (tree is ParserRuleContext context && !(context is VBAParser.WhiteSpaceContext))
            {
                Visit(context);
            }
            return _contextValues;
        }

        public virtual IParseTreeVisitorResults VisitChildren(IRuleNode node)
        {
            if (node is ParserRuleContext context)
            {
                foreach (var child in context.children)
                {
                    Visit(child);
                }
            }
            return _contextValues;
        }

        public virtual IParseTreeVisitorResults VisitTerminal(ITerminalNode node)
        {
            return _contextValues;
        }

        public virtual IParseTreeVisitorResults VisitErrorNode(IErrorNode node)
        {
            return _contextValues;
        }

        private void StoreVisitResult(ParserRuleContext context, IParseTreeValue inspValue)
        {
            OnValueResultCreated(this, new ValueResultEventArgs(context, inspValue));
        }

        private bool ContextHasResult(ParserRuleContext context)
        {
            return _contextValues.Contains(context);
        }

        private void Visit(ParserRuleContext parserRuleContext)
        {
            switch (parserRuleContext)
            {
                case VBAParser.LExprContext lExpr:
                    Visit(lExpr);
                    return;
                case VBAParser.LiteralExprContext litExpr:
                    Visit(litExpr);
                    return;
                case VBAParser.CaseClauseContext caseClause:
                    VisitImpl(caseClause);
                    StoreVisitResult(caseClause, _inspValueFactory.Create(caseClause.GetText()));
                    return;
                case VBAParser.RangeClauseContext rangeClause:
                    VisitImpl(rangeClause);
                    StoreVisitResult(rangeClause, _inspValueFactory.Create(rangeClause.GetText()));
                    return;
                default:
                    if (IsUnaryResultContext(parserRuleContext))
                    {
                        VisitUnaryResultContext(parserRuleContext);
                    }
                    else if (IsBinaryResultContext(parserRuleContext))
                    {
                        VisitBinaryOpEvaluationContext(parserRuleContext);
                    }
                    else if (parserRuleContext is VBAParser.LogicalNotOpContext
                        || parserRuleContext is VBAParser.UnaryMinusOpContext)
                    {
                        VisitUnaryOpEvaluationContext(parserRuleContext);
                    }
                    return;
            }
        }

        private void Visit(VBAParser.LExprContext context)
        {
            if (ContextHasResult(context))
            {
                return;
            }

            IParseTreeValue newResult = null;
            if (TryGetLExprValue(context, out string lexprValue, out string declaredType))
            {
                newResult = _inspValueFactory.Create(lexprValue, declaredType);
            }
            else
            {
                var smplNameExprTypeName = string.Empty;
                var smplName = context.GetDescendent<VBAParser.SimpleNameExprContext>();
                if (TryGetIdentifierReferenceForContext(smplName, out IdentifierReference idRef))
                {
                    var declarationTypeName = GetBaseTypeForDeclaration(idRef.Declaration);
                    newResult = _inspValueFactory.Create(context.GetText(), declarationTypeName);
                }
            }

            if (newResult != null)
            {
                StoreVisitResult(context, newResult);
            }
        }

        private void Visit(VBAParser.LiteralExprContext context)
        {
            if (!ContextHasResult(context))
            {
                var nResult = _inspValueFactory.Create(context.GetText());
                StoreVisitResult(context, nResult);
            }
        }

        private void VisitBinaryOpEvaluationContext(ParserRuleContext context)
        {
            VisitImpl(context);

            RetrieveOpEvaluationElements(context, out (IParseTreeValue LHS, IParseTreeValue RHS, string Symbol) binaryData);
            if(binaryData.LHS is null || binaryData.RHS is null)
            {
                return;
            }

            var calculator = new ParseTreeExpressionEvaluator(_inspValueFactory, context.IsOptionCompareBinary());
            var result = calculator.Evaluate(binaryData.LHS, binaryData.RHS, binaryData.Symbol);

            StoreVisitResult(context, result);
        }

        private void VisitUnaryOpEvaluationContext(ParserRuleContext context)
        {
            VisitImpl(context);
            RetrieveOpEvaluationElements(context, out (IParseTreeValue LHS, IParseTreeValue RHS, string Symbol) unaryData);
            if (unaryData.LHS is null || unaryData.RHS != null)
            {
                return;
            }

            var calculator = new ParseTreeExpressionEvaluator(_inspValueFactory, context.IsOptionCompareBinary());
            var result = calculator.Evaluate(unaryData.LHS, unaryData.Symbol, unaryData.LHS.TypeName);
            StoreVisitResult(context, result);
        }

        private void RetrieveOpEvaluationElements(ParserRuleContext context, out (IParseTreeValue LHS, IParseTreeValue RHS, string Symbol) operandElements)
        {
            operandElements.Symbol = string.Empty;
            operandElements.LHS = null;
            operandElements.RHS = null;
            var values = new List<IParseTreeValue>();
            var contextsOfInterest = NonWhitespaceChildren(context);
            for (var idx = 0; idx < contextsOfInterest.Count(); idx++)
            {
                if (contextsOfInterest.ElementAt(idx) is ParserRuleContext ctxt)
                {
                    if (operandElements.LHS is null)
                    {
                        operandElements.LHS = _contextValues.GetValue(ctxt);
                    }
                    else if (operandElements.RHS is null)
                    {
                        operandElements.RHS = _contextValues.GetValue(ctxt);
                    }
                }
                else
                {
                    operandElements.Symbol = contextsOfInterest.ElementAt(idx).GetText();
                }
            }
        }

        private void VisitUnaryResultContext(ParserRuleContext parserRuleContext)
        {
            VisitImpl(parserRuleContext);

            var contextsOfInterest = ParserRuleContextChildren(parserRuleContext);
            for (var idx = 0; idx < contextsOfInterest.Count(); idx++)
            {
                var ctxt = contextsOfInterest.ElementAt(idx);
                if (_contextValues.Contains(ctxt))
                {
                    var value = _contextValues.GetValue(ctxt);
                    StoreVisitResult(parserRuleContext, value);
                }
            }
        }

        private void VisitImpl(ParserRuleContext context)
        {
            if (!ContextHasResult(context))
            {
                foreach (var ctxt in ParserRuleContextChildren(context))
                {
                    Visit(ctxt);
                }
            }
        }

        private static IEnumerable<ParserRuleContext> ParserRuleContextChildren(IParseTree ptParent)
        {
            return ((ParserRuleContext)ptParent).children.Where(ch => !(ch is VBAParser.WhiteSpaceContext) && ch is ParserRuleContext).Select(item => (ParserRuleContext)item);
        }

        private static IEnumerable<IParseTree> NonWhitespaceChildren(ParserRuleContext ptParent)
        {
            return ptParent.children.Where(ch => !(ch is VBAParser.WhiteSpaceContext));
        }

        private bool TryGetLExprValue(VBAParser.LExprContext lExprContext, out string expressionValue, out string declaredTypeName)
        {
            expressionValue = string.Empty;
            declaredTypeName = string.Empty;
            if (lExprContext.TryGetChildContext(out VBAParser.MemberAccessExprContext memberAccess))
            {
                var member = memberAccess.GetChild<VBAParser.UnrestrictedIdentifierContext>();
                GetContextValue(member, out declaredTypeName, out expressionValue);
                return true;
            }

            if (lExprContext.TryGetChildContext(out VBAParser.SimpleNameExprContext smplName))
            {
                GetContextValue(smplName, out declaredTypeName, out expressionValue);
                return true;
            }

            return false;
        }

        private void GetContextValue(ParserRuleContext context, out string declaredTypeName, out string expressionValue)
        {
            expressionValue = context.GetText();
            declaredTypeName = string.Empty;

            if (TryGetIdentifierReferenceForContext(context, out IdentifierReference rangeClauseIdentifierReference))
            {
                var declaration = rangeClauseIdentifierReference.Declaration;
                expressionValue = rangeClauseIdentifierReference.IdentifierName;
                declaredTypeName = GetBaseTypeForDeclaration(declaration);

                if (declaration.DeclarationType.HasFlag(DeclarationType.Constant)
                    || declaration.DeclarationType.HasFlag(DeclarationType.EnumerationMember))
                {
                    expressionValue = GetConstantDeclarationValueToken(declaration);
                    if (declaration.DeclarationType.HasFlag(DeclarationType.Constant)
                        && declaredTypeName.Equals(Tokens.String))
                    {
                        expressionValue = "\"" + expressionValue + "\"";
                    }
                }
            }
        }

        private bool TryGetIdentifierReferenceForContext(ParserRuleContext context, out IdentifierReference idRef)
        {
            idRef = null;
            if (IdRefRetriever != null)
            {
                (bool success, IdentifierReference idReference) = IdRefRetriever(context);
                idRef = idReference;
                return success;
            }
            else if (State != null) //State is set to non-null for testing
            {
                var identifierReferences = (State.DeclarationFinder.MatchName(context.GetText()).Select(dec => dec.References)).SelectMany(rf => rf)
                    .Where(rf => rf.Context == context);
                if (identifierReferences.Count() == 1)
                {
                    idRef = identifierReferences.First();
                    return true;
                }
            }
            return false;
        }

        private string GetConstantDeclarationValueToken(Declaration constantDeclaration)
        {
            var declarationContextChildren = constantDeclaration.Context.children.ToList();
            var equalsSymbolIndex = declarationContextChildren.FindIndex(ch => ch.Equals(constantDeclaration.Context.GetToken(VBAParser.EQ, 0)));

            var contextsOfInterest = new List<ParserRuleContext>();
            for (int idx = equalsSymbolIndex + 1; idx < declarationContextChildren.Count(); idx++)
            {
                var childCtxt = declarationContextChildren[idx];
                if (!(childCtxt is VBAParser.WhiteSpaceContext))
                {
                    contextsOfInterest.Add((ParserRuleContext)childCtxt);
                }
            }

            foreach (var child in contextsOfInterest)
            {
                Visit(child);
                if (_contextValues.TryGetValue(child, out IParseTreeValue value))
                {
                    return value.ValueText;
                }
            }
            return string.Empty;
        }

        private string GetBaseTypeForDeclaration(Declaration declaration)
        {
            var localDeclaration = declaration;
            var iterationGuard = 0;
            while (!(localDeclaration is null) 
                && !localDeclaration.AsTypeIsBaseType 
                && iterationGuard++ < 5)
            {
                localDeclaration = localDeclaration.AsTypeDeclaration;
            }
            return localDeclaration is null ? declaration.AsTypeName : localDeclaration.AsTypeName;
        }

        private static bool IsUnaryResultContext<T>(T context)
        {
            return context is VBAParser.SelectStartValueContext
                || context is VBAParser.SelectEndValueContext
                || context is VBAParser.ParenthesizedExprContext
                || context is VBAParser.SelectExpressionContext;
        }

        private static bool IsBinaryResultContext<T>(T context)
        {
            if (context is VBAParser.ExpressionContext expressionContext)
            {

                return expressionContext.IsBinaryMathContext()
                    || expressionContext.IsBinaryLogicalContext()
                    || context is VBAParser.ConcatOpContext;
            }
            return false;
        }
    }

    public class ValueResultEventArgs : EventArgs
    {
        public ValueResultEventArgs(ParserRuleContext context, IParseTreeValue value)
        {
            Context = context;
            Value = value;
        }

        public ParserRuleContext Context { set; get; }
        public IParseTreeValue Value { set; get; }
    }
}
