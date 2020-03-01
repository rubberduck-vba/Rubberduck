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
        ICollection<(UnreachableCaseInspection.CaseInspectionResultType resultType, ParserRuleContext context)> InspectForUnreachableCases();
        string SelectExpressionTypeName { get; }
    }

    public class UnreachableCaseInspector : IUnreachableCaseInspector
    {
        private readonly IEnumerable<VBAParser.CaseClauseContext> _caseClauses;
        private readonly ParserRuleContext _caseElseContext;
        private readonly IParseTreeValueFactory _valueFactory;
        private readonly Func<string, ParserRuleContext, string> _getVariableDeclarationTypeName;
        private IParseTreeValue _selectExpressionValue;

        public UnreachableCaseInspector(VBAParser.SelectCaseStmtContext selectCaseContext, 
            IParseTreeVisitorResults inspValues, 
            IParseTreeValueFactory valueFactory,
            Func<string,ParserRuleContext,string> getVariableTypeName = null)
        {
            _valueFactory = valueFactory;
            _caseClauses = selectCaseContext.caseClause();
            _caseElseContext = selectCaseContext.caseElseClause();
            _getVariableDeclarationTypeName = getVariableTypeName;
            ParseTreeValueResults = inspValues;
            SetSelectExpressionTypeName(selectCaseContext, inspValues);
        }

        public List<ParserRuleContext> UnreachableCases { set; get; } = new List<ParserRuleContext>();

        public List<ParserRuleContext> MismatchTypeCases { set; get; } = new List<ParserRuleContext>();

        public List<ParserRuleContext> OverflowCases { set; get; } = new List<ParserRuleContext>();

        public List<ParserRuleContext> InherentlyUnreachableCases { set; get; } = new List<ParserRuleContext>();

        public List<ParserRuleContext> UnreachableCaseElseCases { set; get; } = new List<ParserRuleContext>();

        public ICollection<(UnreachableCaseInspection.CaseInspectionResultType resultType, ParserRuleContext context)>
            AllResults =>
            WithType(UnreachableCaseInspection.CaseInspectionResultType.Unreachable, UnreachableCases)
                .Concat(WithType(UnreachableCaseInspection.CaseInspectionResultType.InherentlyUnreachable, InherentlyUnreachableCases))
                .Concat(WithType(UnreachableCaseInspection.CaseInspectionResultType.MismatchType, MismatchTypeCases))
                .Concat(WithType(UnreachableCaseInspection.CaseInspectionResultType.Overflow, OverflowCases))
                .Concat(WithType(UnreachableCaseInspection.CaseInspectionResultType.CaseElse, UnreachableCaseElseCases))
                .ToList();

        private static IEnumerable<(UnreachableCaseInspection.CaseInspectionResultType type, ParserRuleContext context)>
            WithType(UnreachableCaseInspection.CaseInspectionResultType type, IEnumerable<ParserRuleContext> source)
        {
            return source.Select(context => (type, context));
        }

        public string SelectExpressionTypeName { private set; get; } = string.Empty;

        private IParseTreeVisitorResults ParseTreeValueResults { get; }

        public ICollection<(UnreachableCaseInspection.CaseInspectionResultType resultType, ParserRuleContext context)> InspectForUnreachableCases()
        {
            if (!InspectableTypes.Contains(SelectExpressionTypeName))
            {
                return new List<(UnreachableCaseInspection.CaseInspectionResultType resultType, ParserRuleContext context)>();
            }

            var remainingCasesToInspect = new List<VBAParser.CaseClauseContext>();

            foreach (var caseClause in _caseClauses)
            {
                var containsMismatch = false;
                var containsOverflow = false;
                foreach ( var range in caseClause.rangeClause())
                {
                    var childResults = ParseTreeValueResults.GetChildResults(range);
                    var childValues = childResults
                        .Select(ch => ParseTreeValueResults.GetValue(ch))
                        .ToList();
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

            return AllResults;
        }

        private IExpressionFilter BuildRangeClauseFilter(IEnumerable<VBAParser.CaseClauseContext> caseClauses)
        {
            var rangeClauseFilter = ExpressionFilterFactory.Create(SelectExpressionTypeName);

            if (!(_getVariableDeclarationTypeName is null))
            {
                foreach (var caseClause in caseClauses)
                {
                    foreach (var rangeClause in caseClause.rangeClause())
                    {
                        var expression = GetRangeClauseExpression(rangeClause);
                        if (!expression?.LHS?.ParsesToConstantValue ?? false)
                        {
                            var typeName = _getVariableDeclarationTypeName(expression.LHS.Token, rangeClause);
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
            if (TryDetectTypeHint(selectStmt.selectExpression().GetText(), out var typeName)
                && InspectableTypes.Contains(typeName))
            {
                SelectExpressionTypeName = typeName;
            }
            else if (inspValues.TryGetValue(selectStmt.selectExpression(), out var result)
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
                    if (TryDetectTypeHint(range.GetText(), out var hintTypeName))
                    {
                        caseClauseTypeNames.Add(hintTypeName);
                    }
                    else
                    {
                        var typeNames = range.children
                            .OfType<ParserRuleContext>()
                            .Where(IsResultContext)
                            .Select(inspValues.GetValueType);

                        caseClauseTypeNames.AddRange(typeNames);
                        caseClauseTypeNames.RemoveAll(tp => !InspectableTypes.Contains(tp));
                    }
                }
            }

            if (TryGetSelectExpressionTypeNameFromTypes(caseClauseTypeNames, out var evalTypeName))
            {
                return evalTypeName;
            }

            return string.Empty;
        }

        private static bool TryGetSelectExpressionTypeNameFromTypes(ICollection<string> typeNames, out string typeName)
        {
            typeName = string.Empty;
            if (!typeNames.Any())
            {
                return false;
            }

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

            //Mix of Integer types and rational number types will be evaluated using Double or Currency
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

            if (SymbolList.TypeHintToTypeName.Keys.Any(content.EndsWith))
            {
                var lastChar = content.Substring(content.Length - 1);
                typeName = SymbolList.TypeHintToTypeName[lastChar];
                return true;
            }
            return false;
        }

        private IRangeClauseExpression GetRangeClauseExpression(VBAParser.RangeClauseContext rangeClause)
        {
            var resultContexts = rangeClause.children
                .OfType<ParserRuleContext>()
                .Where(IsResultContext)
                .ToList();

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

            if (rangeClause.IS() != null)
            {
                var isClauseValue = ParseTreeValueResults.GetValue(resultContexts.First());
                var opSymbol = rangeClause.GetChild<VBAParser.ComparisonOperatorContext>().GetText();
                return new IsClauseExpression(isClauseValue, opSymbol);
            }

            if (!TryGetLogicSymbol(resultContexts.First(), out string symbol))
            {
                return new ValueExpression(ParseTreeValueResults.GetValue(resultContexts.First()));
            }

            var resultContext = resultContexts.First();
            var clauseValue = ParseTreeValueResults.GetValue(resultContext);
            if (clauseValue.ParsesToConstantValue)
            {
                return new ValueExpression(clauseValue);
            }

            switch (resultContext)
            {
                case VBAParser.LogicalNotOpContext _:
                    return new UnaryExpression(clauseValue, symbol);
                case VBAParser.RelationalOpContext _:
                case VBAParser.LogicalEqvOpContext _:
                case VBAParser.LogicalImpOpContext _:
                {
                    var (lhs, rhs) = CreateLogicPair(clauseValue, symbol, _valueFactory);
                    if (symbol.Equals(Tokens.Like))
                    {
                        return new LikeExpression(lhs, rhs);
                    }
                    return new BinaryExpression(lhs, rhs, symbol);
                }
                default:
                    return null;
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

        private static (IParseTreeValue lhs, IParseTreeValue rhs) CreateLogicPair(IParseTreeValue value, string opSymbol, IParseTreeValueFactory factory)
        {
            var operands = value.Token.Split(new [] { opSymbol }, StringSplitOptions.None);
            if (operands.Length == 2)
            {
                var lhs = factory.Create(operands[0].Trim());
                var rhs = factory.Create(operands[1].Trim());
                if (opSymbol.Equals(Tokens.Like))
                {
                    rhs = factory.CreateDeclaredType($"\"{rhs.Token}\"", Tokens.String);
                }
                return (lhs, rhs);
            }

            if (operands.Length == 1)
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

        private static readonly IReadOnlyList<string> InspectableTypes = new List<string>
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
