using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Refactorings;
using Rubberduck.VBEditor;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete.UnreachableCaseEvaluation
{
    internal interface IUnreachableCaseInspector
    {
        ICollection<(UnreachableCaseInspection.CaseInspectionResultType resultType, ParserRuleContext context)> InspectForUnreachableCases(
            QualifiedModuleName module, 
            VBAParser.SelectCaseStmtContext selectCaseContext, 
            IParseTreeVisitorResults parseTreeValues,
            DeclarationFinder finder);
        string SelectExpressionTypeName(
            VBAParser.SelectCaseStmtContext selectCaseContext, 
            IParseTreeVisitorResults parseTreeValues);
    }

    internal class UnreachableCaseInspector : IUnreachableCaseInspector
    {
        private readonly IParseTreeValueFactory _valueFactory;

        public UnreachableCaseInspector(
            IParseTreeValueFactory valueFactory)
        {
            _valueFactory = valueFactory;
        }

        public ICollection<(UnreachableCaseInspection.CaseInspectionResultType resultType, ParserRuleContext context)> InspectForUnreachableCases(
            QualifiedModuleName module,
            VBAParser.SelectCaseStmtContext selectCaseContext,
            IParseTreeVisitorResults parseTreeValues,
            DeclarationFinder finder)
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

            var rangeClauseFilter = BuildRangeClauseFilter(module, remainingCasesToInspect, selectExpressionTypeName, parseTreeValues, finder);
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
            var usableClauses = rangeClauseExpressions.Where(expr => expr != null).ToList();
            if (usableClauses.Any(expr => expr.IsMismatch))
            {
                return UnreachableCaseInspection.CaseInspectionResultType.MismatchType;
            }

            if (usableClauses.Any(expr => expr.IsOverflow))
            {
                return UnreachableCaseInspection.CaseInspectionResultType.Overflow;
            }

            if (usableClauses.All(expr => expr.IsInherentlyUnreachable))
            {
                return UnreachableCaseInspection.CaseInspectionResultType.InherentlyUnreachable;
            }

            if (usableClauses.All(expr =>
                expr.IsUnreachable || expr.IsMismatch|| expr.IsOverflow || expr.IsInherentlyUnreachable))
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

        private IExpressionFilter BuildRangeClauseFilter(QualifiedModuleName module, IEnumerable<VBAParser.CaseClauseContext> caseClauses, string selectExpressionTypeName, IParseTreeVisitorResults parseTreeValues, DeclarationFinder finder)
        {
            var rangeClauseFilter = ExpressionFilterFactory.Create(selectExpressionTypeName);

            var rangeClauses = caseClauses.SelectMany(caseClause => caseClause.rangeClause());
            foreach (var rangeClause in rangeClauses)
            {
                var expression = GetRangeClauseExpression(rangeClause, parseTreeValues);
                if (!expression?.LHS?.ParsesToConstantValue ?? false)
                {
                    var typeName = GetVariableTypeName(module, expression.LHS.Token, rangeClause, finder);
                    rangeClauseFilter.AddComparablePredicateFilter(expression.LHS.Token, typeName);
                }
            }

            return rangeClauseFilter;
        }

        private string GetVariableTypeName(QualifiedModuleName module, string variableName, ParserRuleContext ancestor, DeclarationFinder finder)
        {
            if (ancestor == null)
            {
                return string.Empty;
            }

            var descendents = ancestor.GetDescendents<VBAParser.SimpleNameExprContext>()
                .Where(desc => desc.GetText().Equals(variableName))
                .ToList();
            if (!descendents.Any())
            {
                return string.Empty;
            }

            var firstDescendent = descendents.First();
            var (success, reference) = GetIdentifierReferenceForContext(module, firstDescendent, finder);
            return success ?
                GetBaseTypeForDeclaration(reference.Declaration)
                : string.Empty;
        }

        private static (bool success, IdentifierReference idRef) GetIdentifierReferenceForContext(QualifiedModuleName module, ParserRuleContext context, DeclarationFinder finder)
        {
            if (context == null)
            {
                return (false, null);
            }

            var qualifiedSelection = new QualifiedSelection(module, context.GetSelection());

            var identifierReferences =
                finder
                    .IdentifierReferences(qualifiedSelection)
                    .Where(reference => reference.Context == context)
                    .ToList();

            return identifierReferences.Count == 1
                ? (true, identifierReferences.First())
                : (false, null);
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

        public string SelectExpressionTypeName(
            VBAParser.SelectCaseStmtContext selectStmt,
            IParseTreeVisitorResults parseTreeValues)
        {
            var (typeName, _) = SelectExpressionTypeNameAndValue(selectStmt, parseTreeValues);
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
