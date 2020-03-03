using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IUnreachableCaseInspector
    {
        ICollection<(UnreachableCaseInspection.CaseInspectionResultType resultType, ParserRuleContext context)> InspectForUnreachableCases(
            QualifiedModuleName module, 
            VBAParser.SelectCaseStmtContext selectCaseContext, 
            IParseTreeVisitorResults parseTreeValues);
        string SelectExpressionTypeName(
            VBAParser.SelectCaseStmtContext selectCaseContext, 
            IParseTreeVisitorResults parseTreeValues);
    }

    public class UnreachableCaseInspector : IUnreachableCaseInspector
    {
        private readonly IParseTreeValueFactory _valueFactory;
        private readonly Func<string, QualifiedModuleName, ParserRuleContext, string> _getVariableDeclarationTypeName;

        public UnreachableCaseInspector(
            IParseTreeValueFactory valueFactory,
            Func<string, QualifiedModuleName, ParserRuleContext, string> getVariableTypeName = null)
        {
            _valueFactory = valueFactory;
            _getVariableDeclarationTypeName = getVariableTypeName;
        }

        public ICollection<(UnreachableCaseInspection.CaseInspectionResultType resultType, ParserRuleContext context)> InspectForUnreachableCases(
            QualifiedModuleName module,
            VBAParser.SelectCaseStmtContext selectCaseContext,
            IParseTreeVisitorResults parseTreeValues)
        {
            var (selectExpressionTypeName, selectExpressionValue) = SelectExpressionTypeNameAndValue(selectCaseContext, parseTreeValues);

            if (!InspectableTypes.Contains(selectExpressionTypeName))
            {
                return new List<(UnreachableCaseInspection.CaseInspectionResultType resultType, ParserRuleContext context)>();
            }

            var results = new List<(UnreachableCaseInspection.CaseInspectionResultType resultType, ParserRuleContext context)>();
            
            var caseClausesWithInvalidTypeMarker = selectCaseContext.caseClause()
                .Select(caseClause => WithInvalidValueType(caseClause, parseTreeValues))
                .ToList();

            var invalidCaseClausesWithInvalidTypeMarker = caseClausesWithInvalidTypeMarker
                .Where(tpl => tpl.invalidValueType.HasValue)
                .Select(tpl => (tpl.invalidValueType.Value, (ParserRuleContext)tpl.caseClause));

            results.AddRange(invalidCaseClausesWithInvalidTypeMarker);

            var remainingCasesToInspect = caseClausesWithInvalidTypeMarker
                .Where(tpl => tpl.invalidValueType == null)
                .Select(tpl => tpl.caseClause)
                .ToList();

            var rangeClauseFilter = BuildRangeClauseFilter(module, remainingCasesToInspect, selectExpressionTypeName, parseTreeValues);
            if (!(selectExpressionValue is null) && selectExpressionValue.ParsesToConstantValue)
            {
                rangeClauseFilter.SelectExpressionValue = selectExpressionValue;
            }

            foreach (var caseClause in remainingCasesToInspect)
            {
                var rangeClauseExpressions = caseClause.rangeClause()
                    .Select(range =>  GetRangeClauseExpression(range, parseTreeValues))
                    .ToList();

                foreach (var expression in rangeClauseExpressions)
                {
                    rangeClauseFilter.CheckAndAddExpression(expression);
                }

                var invalidRangeType = InvalidRangeExpressionsType(rangeClauseExpressions);
                if (invalidRangeType.HasValue)
                {
                    results.Add((invalidRangeType.Value, caseClause));
                }
            }

            var caseElseClause = selectCaseContext.caseElseClause();
            if (caseElseClause != null && rangeClauseFilter.FiltersAllValues)
            {
                results.Add((UnreachableCaseInspection.CaseInspectionResultType.CaseElse, caseElseClause));
            }

            return results;
        }

        private UnreachableCaseInspection.CaseInspectionResultType? InvalidRangeExpressionsType(ICollection<IRangeClauseExpression> rangeClauseExpressions)
        {
            if (rangeClauseExpressions.Any(expr => expr.IsMismatch))
            {
                return UnreachableCaseInspection.CaseInspectionResultType.MismatchType;
            }

            if (rangeClauseExpressions.Any(expr => expr.IsOverflow))
            {
                return UnreachableCaseInspection.CaseInspectionResultType.Overflow;
            }

            if (rangeClauseExpressions.All(expr => expr.IsInherentlyUnreachable))
            {
                return UnreachableCaseInspection.CaseInspectionResultType.InherentlyUnreachable;
            }

            if (rangeClauseExpressions.All(expr =>
                expr.IsUnreachable || expr.IsMismatch || expr.IsOverflow || expr.IsInherentlyUnreachable))
            {
                return UnreachableCaseInspection.CaseInspectionResultType.Unreachable;
            }

            return null;
        }

        private (UnreachableCaseInspection.CaseInspectionResultType? invalidValueType, VBAParser.CaseClauseContext caseClause) WithInvalidValueType(VBAParser.CaseClauseContext caseClause, IParseTreeVisitorResults parseTreeValues)
        {
            return (InvalidValueType(caseClause, parseTreeValues), caseClause);
        }

        private UnreachableCaseInspection.CaseInspectionResultType? InvalidValueType(VBAParser.CaseClauseContext caseClause, IParseTreeVisitorResults parseTreeValues)
        {
            var rangeClauseChildValues = caseClause
                .rangeClause()
                .SelectMany(parseTreeValues.GetChildResults)
                .Select(parseTreeValues.GetValue)
                .ToList();

            if (rangeClauseChildValues.Any(value => value.IsMismatchExpression))
            {
                return UnreachableCaseInspection.CaseInspectionResultType.MismatchType;
            }

            if (rangeClauseChildValues.Any(value => value.IsOverflowExpression))
            {
                return UnreachableCaseInspection.CaseInspectionResultType.Overflow;
            }

            return null;
        }

        private IExpressionFilter BuildRangeClauseFilter(QualifiedModuleName module, IEnumerable<VBAParser.CaseClauseContext> caseClauses, string selectExpressionTypeName, IParseTreeVisitorResults parseTreeValues)
        {
            var rangeClauseFilter = ExpressionFilterFactory.Create(selectExpressionTypeName);

            if (_getVariableDeclarationTypeName is null)
            {
                return rangeClauseFilter;
            }

            var rangeClauses = caseClauses.SelectMany(caseClause => caseClause.rangeClause());
            foreach (var rangeClause in rangeClauses)
            {
                var expression = GetRangeClauseExpression(rangeClause, parseTreeValues);
                if (!expression?.LHS?.ParsesToConstantValue ?? false)
                {
                    var typeName = _getVariableDeclarationTypeName(expression.LHS.Token, module, rangeClause);
                    rangeClauseFilter.AddComparablePredicateFilter(expression.LHS.Token, typeName);
                }
            }

            return rangeClauseFilter;
        }

        public string SelectExpressionTypeName(
            VBAParser.SelectCaseStmtContext selectStmt,
            IParseTreeVisitorResults parseTreeValues)
        {
            var (typeName, value) = SelectExpressionTypeNameAndValue(selectStmt, parseTreeValues);
            return typeName;
        }

        private (string typeName, IParseTreeValue value) SelectExpressionTypeNameAndValue(
            VBAParser.SelectCaseStmtContext selectStmt,
            IParseTreeVisitorResults parseTreeValues)
        {
            if (TryDetectTypeHint(selectStmt.selectExpression().GetText(), out var typeName)
                && InspectableTypes.Contains(typeName))
            {
                return (typeName, null);
            }

            if (parseTreeValues.TryGetValue(selectStmt.selectExpression(), out var result)
                && InspectableTypes.Contains(result.ValueType))
            {
                return (result.ValueType, result);
            }

            return (DeriveTypeFromCaseClauses(parseTreeValues, selectStmt), null);
        }

        private string DeriveTypeFromCaseClauses(IParseTreeVisitorResults parseTreeValues, VBAParser.SelectCaseStmtContext selectStmt)
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
                            .Select(parseTreeValues.GetValueType);

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

        private IRangeClauseExpression GetRangeClauseExpression(VBAParser.RangeClauseContext rangeClause, IParseTreeVisitorResults parseTreeValues)
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
                var rangeStartValue = parseTreeValues.GetValue(rangeClause.GetChild<VBAParser.SelectStartValueContext>());
                var rangeEndValue = parseTreeValues.GetValue(rangeClause.GetChild<VBAParser.SelectEndValueContext>());
                return new RangeOfValuesExpression((rangeStartValue, rangeEndValue));
            }

            if (rangeClause.IS() != null)
            {
                var isClauseValue = parseTreeValues.GetValue(resultContexts.First());
                var opSymbol = rangeClause.GetChild<VBAParser.ComparisonOperatorContext>().GetText();
                return new IsClauseExpression(isClauseValue, opSymbol);
            }

            if (!TryGetLogicSymbol(resultContexts.First(), out string symbol))
            {
                return new ValueExpression(parseTreeValues.GetValue(resultContexts.First()));
            }

            var resultContext = resultContexts.First();
            var clauseValue = parseTreeValues.GetValue(resultContext);
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
