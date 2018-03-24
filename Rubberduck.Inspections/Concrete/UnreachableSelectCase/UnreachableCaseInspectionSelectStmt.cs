using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Inspections.Concrete.UnreachableSelectCase
{
    public interface IUnreachableCaseInspectionSelectStmt
    {
        string EvaluationTypeName { get; }
        void InspectForUnreachableCases();
        List<ParserRuleContext> UnreachableCases { get; }
        List<ParserRuleContext> MismatchTypeCases { get; }
        List<ParserRuleContext> UnreachableCaseElseCases { get; }
    }

    public class UnreachableCaseInspectionSelectStmt : UnreachableCaseInspectionContext, IUnreachableCaseInspectionSelectStmt
    {
        private readonly VBAParser.SelectCaseStmtContext _selectCaseContext;
        private List<ParserRuleContext> _unreachableResults;
        private List<ParserRuleContext> _misMatchResults;
        private List<ParserRuleContext> _caseElseResults;
        private string _evalTypeName;

        public UnreachableCaseInspectionSelectStmt(VBAParser.SelectCaseStmtContext selectCaseContext, IUCIValueResults inspValues, IUCIRangeClauseFilterFactory factory, IUCIValueFactory valueFactory)
            : base(selectCaseContext, inspValues, factory, valueFactory)
        {
            _selectCaseContext = selectCaseContext;
            _unreachableResults = new List<ParserRuleContext>();
            _misMatchResults = new List<ParserRuleContext>();
            _caseElseResults = new List<ParserRuleContext>();
            _evalTypeName = null;
        }

        public string EvaluationTypeName
        {
            get => _evalTypeName ?? DetermineSelectCaseEvaluationTypeName(Context, _inspValues, _factoryRangeClauseFilter);
        }

        public List<ParserRuleContext> UnreachableCases => _unreachableResults;
        public List<ParserRuleContext> MismatchTypeCases => _misMatchResults;
        public List<ParserRuleContext> UnreachableCaseElseCases => _caseElseResults;
        protected override bool IsResultContext<TContext>(TContext context)
        {
            return context is VBAParser.SelectExpressionContext;
        }

        public void InspectForUnreachableCases()
        {
            _evalTypeName = DetermineSelectCaseEvaluationTypeName(Context, _inspValues, _factoryRangeClauseFilter);
            if (_evalTypeName is null || _evalTypeName.Equals(Tokens.Variant))
            {
                return;
            }

            var cummulativeRangeFilter = CreateInspectionFilter();
            foreach (var caseClause in _selectCaseContext.caseClause())
            {
                var inspectedRanges = new List<IUnreachableCaseInspectionRange>();
                foreach (var range in caseClause.rangeClause())
                {
                    var inspRange = WrapContext(range);
                    if (inspRange.IsReachable(cummulativeRangeFilter))
                    {
                        cummulativeRangeFilter.Add(inspRange.RangeClause);
                    }
                    inspectedRanges.Add(inspRange);
                }

                if (inspectedRanges.All(ir => ir.HasIncompatibleType))
                {
                    _misMatchResults.Add(caseClause);
                }
                else if (inspectedRanges.All(ir => !ir.HasCoverage))
                {
                    _unreachableResults.Add(caseClause);
                }
            }
            if (cummulativeRangeFilter.FiltersAllValues && !(_selectCaseContext.caseElseClause() is null))
            {
                _caseElseResults.Add(_selectCaseContext.caseElseClause());
            }
        }

        private IUnreachableCaseInspectionRange WrapContext(VBAParser.RangeClauseContext range)
        {
            IUnreachableCaseInspectionRange inspectableRange = new UnreachableCaseInspectionRange(range, _inspValues, _factoryRangeClauseFilter, _factoryValue)
            {
                EvaluationTypeName = EvaluationTypeName
            };
            return inspectableRange;
        }

        private IUCIRangeClauseFilter CreateInspectionFilter()
        {
            return _factoryRangeClauseFilter.Create(EvaluationTypeName, _factoryValue, _factoryRangeClauseFilter);
        }

        private string DetermineSelectCaseEvaluationTypeName(ParserRuleContext context, IUCIValueResults inspValues, IUCIRangeClauseFilterFactory factory)
        {
            var selectStmt = (VBAParser.SelectCaseStmtContext)context;
            if (TryDetectTypeHint(selectStmt.selectExpression().GetText(), out _evalTypeName))
            {
                return _evalTypeName;
            }

            var typeName = string.Empty;
            if (inspValues.TryGetValue(selectStmt.selectExpression(), out IUCIValue result))
            {
                _evalTypeName = result.TypeName;
            }

            if (TypeNameCandBeInspected(_evalTypeName))
            {
                return _evalTypeName;
            }
            return DeriveTypeFromCaseClauses(inspValues, selectStmt, factory);
        }

        private static bool TypeNameCandBeInspected(string typeName) => !(typeName == string.Empty || typeName == Tokens.Variant);

        private string DeriveTypeFromCaseClauses(IUCIValueResults inspValues, VBAParser.SelectCaseStmtContext selectStmt, IUCIRangeClauseFilterFactory factory)
        {
            var caseClauseTypeNames = new List<string>();
            foreach (var caseClause in selectStmt.caseClause())
            {
                foreach (var range in caseClause.rangeClause())
                {
                    if (TryDetectTypeHint(range.GetText(), out string hintTypeName))
                    {
                        caseClauseTypeNames.Add(hintTypeName);
                        continue;
                    }

                    var inspRange = new UnreachableCaseInspectionRange(range, inspValues, factory, _factoryValue);
                    if (inspRange.IsValueRange)
                    {
                        caseClauseTypeNames.Add(inspValues.GetTypeName(inspRange.RangeStart));
                        caseClauseTypeNames.Add(inspValues.GetTypeName(inspRange.RangeEnd));
                    }
                    else if (inspRange.IsRelationalOp)
                    {
                        caseClauseTypeNames.Add(Tokens.Boolean);
                    }
                    else
                    {
                        caseClauseTypeNames.Add(inspValues.GetTypeName(inspRange.Result));
                    }
                }
            }

            if (TryDetermineEvaluationTypeFromTypes(caseClauseTypeNames, out _evalTypeName))
            {
                return _evalTypeName;
            }
            return string.Empty;
        }

        private static bool TryDetectTypeHint(string content, out string typeName)
        {
            typeName = string.Empty;
            if (SymbolList.TypeHintToTypeName.Keys.Any(th => content.EndsWith(th)))
            {
                var lastChar = content.Substring(content.Length - 1);
                typeName = SymbolList.TypeHintToTypeName[lastChar];
                return true;
            }
            return false;
        }
    }
}
