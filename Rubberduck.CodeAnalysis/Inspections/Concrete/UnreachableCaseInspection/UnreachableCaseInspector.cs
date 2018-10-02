using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IUnreachableCaseInspector
    {
        void InspectForUnreachableCases();
        string SelectExpressionTypeName { get; }
        Func<string, ParserRuleContext, string> GetVariableDeclarationTypeName { set; get; }
        List<ParserRuleContext> UnreachableCases { get; }
        List<ParserRuleContext> InherentlyUnreachableCases { get; }
        List<ParserRuleContext> MismatchTypeCases { get; }
        List<ParserRuleContext> OverflowCases { get; }
        List<ParserRuleContext> UnreachableCaseElseCases { get; }
    }

    public class UnreachableCaseInspector : IUnreachableCaseInspector
    {
        private readonly IEnumerable<VBAParser.CaseClauseContext> _caseClauses;
        private readonly ParserRuleContext _caseElseContext;
        private readonly IParseTreeValueFactory _valueFactory;
        private IParseTreeValue _selectExpressionValue;

        public UnreachableCaseInspector(VBAParser.SelectCaseStmtContext selectCaseContext, 
            IParseTreeVisitorResults inspValues, 
            IParseTreeValueFactory valueFactory,
            Func<string,ParserRuleContext,string> GetVariableTypeName = null)
        {
            _valueFactory = valueFactory;
            _caseClauses = selectCaseContext.caseClause();
            _caseElseContext = selectCaseContext.caseElseClause();
            GetVariableDeclarationTypeName = GetVariableTypeName;
            ParseTreeValueResults = inspValues;
            SetSelectExpressionTypeName(selectCaseContext as ParserRuleContext, inspValues);
        }

        public Func<string, ParserRuleContext, string> GetVariableDeclarationTypeName { set; get; }

        public List<ParserRuleContext> UnreachableCases { set; get; } = new List<ParserRuleContext>();

        public List<ParserRuleContext> MismatchTypeCases { set; get; } = new List<ParserRuleContext>();

        public List<ParserRuleContext> OverflowCases { set; get; } = new List<ParserRuleContext>();

        public List<ParserRuleContext> InherentlyUnreachableCases { set; get; } = new List<ParserRuleContext>();

        public List<ParserRuleContext> UnreachableCaseElseCases { set; get; } = new List<ParserRuleContext>();

        public string SelectExpressionTypeName { private set; get; } = string.Empty;

        private IParseTreeVisitorResults ParseTreeValueResults { set; get; }

        public void InspectForUnreachableCases()
        {
            if (!InspectableTypes.Contains(SelectExpressionTypeName))
            {
                return;
            }

            var remainingCasesToInspect = new List<VBAParser.CaseClauseContext>();

            foreach (var caseClause in _caseClauses)
            {
                var containsMismatch = false;
                var containsOverflow = false;
                foreach ( var range in caseClause.rangeClause())
                {
                    var childResults = ParseTreeValueResults.GetChildResults(range);
                    var childValues = childResults.Select(ch => ParseTreeValueResults.GetValue(ch));
                    if (childValues.Any(chr => chr.IsMismatchExpression))
                    {
                        containsMismatch = true;
                    }
                    if (childValues.Any(chr => chr.IsOverflowExpression))
                    {
                        containsOverflow = true;
                    }
                }
                if (containsMismatch)
                {
                    MismatchTypeCases.Add(caseClause);
                }
                else if (containsOverflow)
                {
                    OverflowCases.Add(caseClause);
                }
                else
                {
                    remainingCasesToInspect.Add(caseClause);
                }
            }

            var rangeClauseFilter = BuildRangeClauseFilter(remainingCasesToInspect);
            if (!(_selectExpressionValue is null) && _selectExpressionValue.ParsesToConstantValue)
            {
                rangeClauseFilter.SelectExpressionValue = _selectExpressionValue;
            }

            foreach (var caseClause in remainingCasesToInspect)
            {
                var rangeClauseExpressions = (from range in caseClause.rangeClause()
                                              select GetRangeClauseExpression(range)).ToList();

                rangeClauseExpressions.ForEach(expr => rangeClauseFilter.AddExpression(expr));

                if (rangeClauseExpressions.Any(expr => expr.IsMismatch))
                {
                    MismatchTypeCases.Add(caseClause);
                }
                else if (rangeClauseExpressions.Any(expr => expr.IsOverflow))
                {
                    OverflowCases.Add(caseClause);
                }
                else if (rangeClauseExpressions.All(expr => expr.IsInherentlyUnreachable))
                {
                    InherentlyUnreachableCases.Add(caseClause);
                }
                else if (rangeClauseExpressions.All(expr => expr.IsUnreachable || expr.IsMismatch || expr.IsOverflow || expr.IsInherentlyUnreachable))
                {
                    UnreachableCases.Add(caseClause);
                }
            }

            if (_caseElseContext != null && rangeClauseFilter.FiltersAllValues)
            {
                UnreachableCaseElseCases.Add(_caseElseContext);
            }
        }

        private IExpressionFilter BuildRangeClauseFilter(IEnumerable<VBAParser.CaseClauseContext> caseClauses)
        {
            var rangeClauseFilter = ExpressionFilterFactory.Create(SelectExpressionTypeName);

            if (!(GetVariableDeclarationTypeName is null))
            {
                foreach (var caseClause in caseClauses)
                {
                    foreach (var rangeClause in caseClause.rangeClause())
                    {
                        var expression = GetRangeClauseExpression(rangeClause);
                        if (!expression.LHS.ParsesToConstantValue)
                        {
                            var typeName = GetVariableDeclarationTypeName(expression.LHS.Token, rangeClause);
                            rangeClauseFilter.AddComparablePredicateFilter(expression.LHS.Token, typeName);
                        }
                    }
                }
            }
            return rangeClauseFilter;
        }

        private void SetSelectExpressionTypeName(ParserRuleContext context, IParseTreeVisitorResults inspValues)
        {
            var selectStmt = (VBAParser.SelectCaseStmtContext)context;
            if (TryDetectTypeHint(selectStmt.selectExpression().GetText(), out string typeName)
                && InspectableTypes.Contains(typeName))
            {
                SelectExpressionTypeName = typeName;
            }
            else if (inspValues.TryGetValue(selectStmt.selectExpression(), out IParseTreeValue result)
                && InspectableTypes.Contains(result.ValueType))
            {
                _selectExpressionValue = result;
                SelectExpressionTypeName = result.ValueType;
            }
            else
            {
                SelectExpressionTypeName = DeriveTypeFromCaseClauses(inspValues, selectStmt);
            }
        }

        private string DeriveTypeFromCaseClauses(IParseTreeVisitorResults inspValues, VBAParser.SelectCaseStmtContext selectStmt)
        {
            var caseClauseTypeNames = new List<string>();
            foreach (var caseClause in selectStmt.caseClause())
            {
                foreach (var range in caseClause.rangeClause())
                {
                    if (TryDetectTypeHint(range.GetText(), out string hintTypeName))
                    {
                        caseClauseTypeNames.Add(hintTypeName);
                    }
                    else
                    {
                        var typeNames = from context in range.children
                                where context is ParserRuleContext 
                                    && IsResultContext(context)
                                select inspValues.GetValueType(context as ParserRuleContext);

                        caseClauseTypeNames.AddRange(typeNames);
                        caseClauseTypeNames.RemoveAll(tp => !InspectableTypes.Contains(tp));
                    }
                }
            }

            if (TryGetSelectExpressionTypeNameFromTypes(caseClauseTypeNames, out string evalTypeName))
            {
                return evalTypeName;
            }
            return string.Empty;
        }

        private static bool TryGetSelectExpressionTypeNameFromTypes(IEnumerable<string> typeNames, out string typeName)
        {
            typeName = string.Empty;
            if (!typeNames.Any()) { return false; }

            //If everything is declared as a Variant , we do not attempt to inspect the selectStatement
            if (typeNames.All(tn => tn.Equals(Tokens.Variant)))
            {
                return false;
            }

            //If all match, the typeName is easy...This is the only way to return "String" or "Date".
            if (typeNames.All(tn => new List<string>() { typeNames.First() }.Contains(tn)))
            {
                typeName = typeNames.First();
                return true;
            }
            //Integral numbers will be evaluated using Long
            if (typeNames.All(tn => new List<string>() { Tokens.Long, Tokens.Integer, Tokens.Byte }.Contains(tn)))
            {
                typeName = Tokens.Long;
                return true;
            }

            //Mix of Integertypes and rational number types will be evaluated using Double or Currency
            if (typeNames.All(tn => new List<string>() { Tokens.Long, Tokens.Integer, Tokens.Byte, Tokens.Single, Tokens.Double, Tokens.Currency }.Contains(tn)))
            {
                typeName = typeNames.Any(tk => tk.Equals(Tokens.Currency)) ? Tokens.Currency : Tokens.Double;
                return true;
            }
            return false;
        }

        private static bool TryDetectTypeHint(string content, out string typeName)
        {
            typeName = string.Empty;
            if (content.StartsWith("#") && content.EndsWith("#"))
            {
                return false;
            }

            if (SymbolList.TypeHintToTypeName.Keys.Any(th => content.EndsWith(th)))
            {
                var lastChar = content.Substring(content.Length - 1);
                typeName = SymbolList.TypeHintToTypeName[lastChar];
                return true;
            }
            return false;
        }

        private IRangeClauseExpression GetRangeClauseExpression(VBAParser.RangeClauseContext rangeClause)
        {
            var resultContexts = from ctxt in rangeClause.children
                             where ctxt is ParserRuleContext && IsResultContext(ctxt)
                             select ctxt as ParserRuleContext;

            if (!resultContexts.Any())
            {
                return null;
            }

            if (rangeClause.TO() != null)
            {
                var rangeStartValue = ParseTreeValueResults.GetValue(rangeClause.GetChild<VBAParser.SelectStartValueContext>());
                var rangeEndValue = ParseTreeValueResults.GetValue(rangeClause.GetChild<VBAParser.SelectEndValueContext>());
                return new RangeOfValuesExpression((rangeStartValue, rangeEndValue));
            }
            else if (rangeClause.IS() != null)
            {
                var clauseValue = ParseTreeValueResults.GetValue(resultContexts.First());
                var opSymbol = rangeClause.GetChild<VBAParser.ComparisonOperatorContext>().GetText();
                return new IsClauseExpression(clauseValue, opSymbol);
            }
            else if (TryGetLogicSymbol(resultContexts.First(), out string symbol))
            {
                var resultContext = resultContexts.First();
                var clauseValue = ParseTreeValueResults.GetValue(resultContext);
                if (clauseValue.ParsesToConstantValue)
                {
                    return new ValueExpression(clauseValue);
                }

                if (resultContext is VBAParser.LogicalNotOpContext)
                {
                    return new UnaryExpression(clauseValue, symbol);
                }
                else if (resultContext is VBAParser.RelationalOpContext
                        || resultContext is VBAParser.LogicalEqvOpContext
                        || resultContext is VBAParser.LogicalImpOpContext)
                {
                    (IParseTreeValue lhs, IParseTreeValue rhs) = CreateLogicPair(clauseValue, symbol, _valueFactory);
                    if (symbol.Equals(Tokens.Like))
                    {
                        return new LikeExpression(lhs, rhs);
                    }
                    return new BinaryExpression(lhs, rhs, symbol);
                }
                return null;
            }
            else
            {
                return new ValueExpression(ParseTreeValueResults.GetValue(resultContexts.First()));
            }
        }

        private static bool TryGetLogicSymbol(ParserRuleContext context, out string opSymbol)
        {
            opSymbol = string.Empty;
            if (context is VBAParser.ExpressionContext expressionContext)
            {
                return expressionContext.TryGetLogicalContextSymbol(out opSymbol);
            }
            return false;
        }

        private static (IParseTreeValue lhs, IParseTreeValue rhs)
            CreateLogicPair(IParseTreeValue value, string opSymbol, IParseTreeValueFactory factory)
        {
            var operands = value.Token.Split(new string[] { opSymbol }, StringSplitOptions.None);
            if (operands.Count() == 2)
            {
                var lhs = factory.Create(operands[0].Trim());
                var rhs = factory.Create(operands[1].Trim());
                if (opSymbol.Equals(Tokens.Like))
                {
                    rhs = factory.CreateDeclaredType($"\"{rhs.Token}\"", Tokens.String);
                }
                return (lhs, rhs);
            }

            if (operands.Count() == 1)
            {
                var lhs = factory.Create(operands[0].Trim());
                return (lhs, null);
            }
            return (null, null);
        }

        private static bool IsResultContext<TContext>(TContext context)
        {
            if (context is VBAParser.ExpressionContext expressionContext)
            {
                return  expressionContext.IsMathContext()
                        || expressionContext.IsLogicalContext()
                        || expressionContext is VBAParser.ConcatOpContext
                        || expressionContext is VBAParser.ParenthesizedExprContext
                        || expressionContext is VBAParser.LExprContext
                        || expressionContext is VBAParser.LiteralExprContext;
            }
            return context is VBAParser.SelectStartValueContext
                    || context is VBAParser.SelectEndValueContext;
        }

        private static List<string> InspectableTypes = new List<string>()
        {
            Tokens.Byte,
            Tokens.Integer,
            Tokens.Int,
            Tokens.Long,
            Tokens.LongLong,
            Tokens.Single,
            Tokens.Double,
            Tokens.Decimal,
            Tokens.Currency,
            Tokens.Boolean,
            Tokens.Date,
            Tokens.String
        };
    }
}
