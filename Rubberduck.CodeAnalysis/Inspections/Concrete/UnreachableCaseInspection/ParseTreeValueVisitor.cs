using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IParseTreeValueVisitor : IParseTreeVisitor<IParseTreeVisitorResults>
    {
        event EventHandler<ValueResultEventArgs> OnValueResultCreated;
    }

    public interface ITestParseTreeVisitor
    {
        void InjectValuedDeclarationEvaluator(Func<Declaration, (bool, string, string)> func);
    }

    public class ParseTreeValueVisitor : IParseTreeValueVisitor, ITestParseTreeVisitor
    {
        private class EnumMember
        {
            public EnumMember(VBAParser.EnumerationStmt_ConstantContext constContext, long initValue)
            {
                ConstantContext = constContext;
                Value = initValue;
                HasAssignment = constContext.children.Any(ch => ch.Equals(constContext.GetToken(VBAParser.EQ, 0)));
            }
            public VBAParser.EnumerationStmt_ConstantContext ConstantContext { set; get; }
            public long Value { set; get; }
            public bool HasAssignment { set; get; }
        }

        private IParseTreeVisitorResults _contextValues;
        private IParseTreeValueFactory _inspValueFactory;
        private List<VBAParser.EnumerationStmtContext> _enumStmtContexts;
        private List<EnumMember> _enumMembers;

        public ParseTreeValueVisitor(IParseTreeValueFactory valueFactory, List<VBAParser.EnumerationStmtContext> allEnums, Func<ParserRuleContext, (bool success, IdentifierReference idRef)> idRefRetriever)
        {
            _inspValueFactory = valueFactory;
            IdRefRetriever = idRefRetriever;
            _contextValues = new ParseTreeVisitorResults();
            OnValueResultCreated += _contextValues.OnNewValueResult;
            _enumStmtContexts = allEnums;
        }

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

        private bool HasResult(ParserRuleContext context)
         => _contextValues.Contains(context);

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
                    VisitChildren(caseClause);
                    StoreVisitResult(caseClause, _inspValueFactory.Create(caseClause.GetText()));
                    return;
                case VBAParser.RangeClauseContext rangeClause:
                    VisitChildren(rangeClause);
                    StoreVisitResult(rangeClause, _inspValueFactory.Create(rangeClause.GetText()));
                    return;
                default:
                    if (IsUnaryResultContext(parserRuleContext))
                    {
                        VisitUnaryResultContext(parserRuleContext);
                    }
                    else if (IsBinaryOpEvaluationContext(parserRuleContext))
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
            if (HasResult(context))
            {
                return;
            }

            IParseTreeValue newResult = null;
            if (TryGetLExprValue(context, out string lexprValue, out string declaredType))
            {
                newResult = _inspValueFactory.CreateDeclaredType(lexprValue, declaredType);
            }
            else
            {
                var smplNameExprTypeName = string.Empty;
                var smplName = context.GetDescendent<VBAParser.SimpleNameExprContext>();
                if (TryGetIdentifierReferenceForContext(smplName, out IdentifierReference idRef))
                {
                    var declarationTypeName = GetBaseTypeForDeclaration(idRef.Declaration);
                    newResult = _inspValueFactory.CreateDeclaredType(context.GetText(), declarationTypeName);
                }
            }

            if (newResult != null)
            {
                StoreVisitResult(context, newResult);
            }
        }

        private void Visit(VBAParser.LiteralExprContext context)
        {
            if (!HasResult(context))
            {
                var nResult = _inspValueFactory.Create(context.GetText());
                StoreVisitResult(context, nResult);
            }
        }

        private void VisitBinaryOpEvaluationContext(ParserRuleContext context)
        {
            VisitChildren(context);

            RetrieveOpEvaluationElements(context, out (IParseTreeValue LHS, IParseTreeValue RHS, string Symbol) binaryData);
            if (binaryData.LHS is null || binaryData.RHS is null)
            {
                return;
            }
            if (binaryData.LHS.IsOverflowExpression)
            {
                StoreVisitResult(context, binaryData.LHS);
                return;
            }

            if (binaryData.RHS.IsOverflowExpression)
            {
                StoreVisitResult(context, binaryData.RHS);
                return;
            }

            var calculator = new ParseTreeExpressionEvaluator(_inspValueFactory, context.IsOptionCompareBinary());
            var result = calculator.Evaluate(binaryData.LHS, binaryData.RHS, binaryData.Symbol);

            StoreVisitResult(context, result);
        }

        private void VisitUnaryOpEvaluationContext(ParserRuleContext context)
        {
            VisitChildren(context);

            RetrieveOpEvaluationElements(context, out (IParseTreeValue LHS, IParseTreeValue RHS, string Symbol) unaryData);
            if (unaryData.LHS is null || unaryData.RHS != null)
            {
                return;
            }

            var calculator = new ParseTreeExpressionEvaluator(_inspValueFactory, context.IsOptionCompareBinary());
            var result = calculator.Evaluate(unaryData.LHS, unaryData.Symbol);
            StoreVisitResult(context, result);
        }

        private void RetrieveOpEvaluationElements(ParserRuleContext context, out (IParseTreeValue LHS, IParseTreeValue RHS, string Symbol) operandElements)
        {
            operandElements = (null, null, string.Empty);
            foreach (var child in NonWhitespaceChildren(context))
            {
                if (child is ParserRuleContext ctxt)
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
                    operandElements.Symbol = child.GetText();
                }
            }
        }

        private void VisitUnaryResultContext(ParserRuleContext parserRuleContext)
        {
            VisitChildren(parserRuleContext);

            foreach (var ctxt in ParserRuleContextChildren(parserRuleContext).Where(ct => HasResult(ct)))
            {
                StoreVisitResult(parserRuleContext, _contextValues.GetValue(ctxt));
                return;
            }
        }

        private void VisitChildren(ParserRuleContext context)
        {
            if (!HasResult(context))
            {
                foreach (var ctxt in ParserRuleContextChildren(context))
                {
                    Visit(ctxt);
                }
            }
        }

        private static IEnumerable<ParserRuleContext> ParserRuleContextChildren(ParserRuleContext ptParent)
            => NonWhitespaceChildren(ptParent).Where(ch => ch is ParserRuleContext).Cast<ParserRuleContext>();

        private static IEnumerable<IParseTree> NonWhitespaceChildren(ParserRuleContext ptParent)
            => ptParent.children.Where(ch => !(ch is VBAParser.WhiteSpaceContext));

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

            if (lExprContext.TryGetChildContext(out VBAParser.IndexExprContext idxExpr)
                && ParseTreeValue.TryGetNonPrintingControlCharCompareToken(idxExpr.GetText(), out string comparableToken))
            {
                declaredTypeName = Tokens.String;
                expressionValue = comparableToken;
                return true;
            }

            return false;
        }

        private Func<Declaration, (bool, string, string)> _valueDeclarationEvaluator;
        private Func<Declaration, (bool, string, string)> ValuedDeclarationEvaluator
        {
            set
            {
                _valueDeclarationEvaluator = value;
            }
            get
            {
                return _valueDeclarationEvaluator ?? GetValuedDeclaration;
            }
        }


        private (bool IsType, string ExpressionValue, string TypeName) GetValuedDeclaration(Declaration declaration)
        {
            if (declaration is ValuedDeclaration valuedDeclaration)
            {
                var typeName = GetBaseTypeForDeclaration(declaration);
                return (true, valuedDeclaration.Expression, typeName);
            }
            return (false, null, null);
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

                (bool IsValuedDeclaration, string ExpressionValue, string TypeName) = ValuedDeclarationEvaluator(declaration);

                if( IsValuedDeclaration)
                {
                    expressionValue = ExpressionValue;
                    declaredTypeName = TypeName;

                    if (ParseTreeValue.TryGetNonPrintingControlCharCompareToken(expressionValue, out string resolvedValue))
                    {
                        expressionValue = resolvedValue;
                        declaredTypeName = Tokens.String;
                        return;
                    }
                    else if (long.TryParse(expressionValue, out _))
                    {
                        return;
                    }
                }

                if (declaration.DeclarationType.HasFlag(DeclarationType.Constant))
                {
                    expressionValue = GetConstantContextValueToken(declaration.Context);
                }
                else if (declaration.DeclarationType.HasFlag(DeclarationType.EnumerationMember))
                {
                    declaredTypeName = Tokens.Long;
                    expressionValue = GetConstantContextValueToken(declaration.Context);
                    if (expressionValue.Equals(string.Empty))
                    {
                        if (_enumMembers is null)
                        {
                            LoadEnumMemberValues();
                        }
                        var enumValue = _enumMembers.SingleOrDefault(dt => dt.ConstantContext == declaration.Context);
                        expressionValue = enumValue?.Value.ToString() ?? string.Empty;
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
            return false;
        }

        private string GetConstantContextValueToken(ParserRuleContext context)
        {
            if (context is null)
            {
                return string.Empty;
            }

            var declarationContextChildren = context.children.ToList();
            var equalsSymbolIndex = declarationContextChildren.FindIndex(ch => ch.Equals(context.GetToken(VBAParser.EQ, 0)));

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
                    return value.Token;
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

        private static bool IsBinaryOpEvaluationContext<T>(T context)
        {
            if (context is VBAParser.ExpressionContext expressionContext)
            {

                return expressionContext.IsBinaryMathContext()
                    || expressionContext.IsBinaryLogicalContext()
                    || context is VBAParser.ConcatOpContext;
            }
            return false;
        }

        public void InjectValuedDeclarationEvaluator( Func<Declaration, (bool, string, string)> func)
            => ValuedDeclarationEvaluator = func;

        private void LoadEnumMemberValues()
        {
            _enumMembers = new List<EnumMember>();
            foreach (var enumStmt in _enumStmtContexts)
            {
                long enumAssignedValue = -1;
                var enumConstContexts = enumStmt.children.Where(ch => ch is VBAParser.EnumerationStmt_ConstantContext).Cast<VBAParser.EnumerationStmt_ConstantContext>();
                foreach (var enumConstContext in enumConstContexts)
                {
                    enumAssignedValue++;
                    var enumMember = new EnumMember(enumConstContext, enumAssignedValue);
                    if (enumMember.HasAssignment)
                    {
                        Visit(enumMember.ConstantContext);

                        var valueText = GetConstantContextValueToken(enumMember.ConstantContext);
                        if (!valueText.Equals(string.Empty))
                        {
                            enumMember.Value = long.Parse(valueText);
                            enumAssignedValue = enumMember.Value;
                        }
                    }
                    _enumMembers.Add(enumMember);
                }
            }
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
