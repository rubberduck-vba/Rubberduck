using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IUnreachableCaseInspectionSelectStmt
    {
        string EvaluationTypeName { get; }
        void InspectForUnreachableCases();
        List<ParserRuleContext> UnreachableCases { get; }
        List<ParserRuleContext> MismatchTypeCases { get; }
        List<ParserRuleContext> UnreachableCaseElseCases { get; }
        ParserRuleContext Context { get; }
    }

    public class UnreachableCaseInspectionSelectStmt : UnreachableCaseInspectionContext, IUnreachableCaseInspectionSelectStmt
    {
        private readonly VBAParser.SelectCaseStmtContext _selectCaseContext;
        private IUnreachableCaseInspectionRangeFactory _inspectionRangeFactory;
        private List<ParserRuleContext> _unreachableResults;
        private List<ParserRuleContext> _misMatchResults;
        private List<ParserRuleContext> _caseElseResults;
        private string _evalTypeName;

        public UnreachableCaseInspectionSelectStmt(VBAParser.SelectCaseStmtContext selectCaseContext, IUCIValueResults inspValues, IUnreachableCaseInspectionFactoryFactory factoryFactory)
            : base(selectCaseContext, inspValues, factoryFactory)
        {
            _selectCaseContext = selectCaseContext;
            _unreachableResults = new List<ParserRuleContext>();
            _misMatchResults = new List<ParserRuleContext>();
            _caseElseResults = new List<ParserRuleContext>();
            _evalTypeName = null;
            _inspectionRangeFactory = factoryFactory.CreateUnreachableCaseInspectionRangeFactory();
        }

        public string EvaluationTypeName => _evalTypeName 
            ?? DetermineSelectCaseEvaluationTypeName(Context, ParseTreeValueResults, FilterFactory);

        public void InspectForUnreachableCases()
        {
            if (!InspectionCanEvaluateTypeName(EvaluationTypeName))
            {
                return;
            }

            var cummulativeRangeFilter = FilterFactory.Create(EvaluationTypeName, ValueFactory);
            foreach (var caseClause in _selectCaseContext.caseClause())
            {
                var inspectedRanges = new List<IUnreachableCaseInspectionRange>();
                foreach (var range in caseClause.rangeClause())
                {
                    var inspectionRange = _inspectionRangeFactory.Create(EvaluationTypeName, range, _inspValues);
                    if (inspectionRange.IsReachable(cummulativeRangeFilter))
                    {
                        cummulativeRangeFilter.Add(inspectionRange.AsFilter);
                    }
                    inspectedRanges.Add(inspectionRange);
                }

                if (inspectedRanges.All(ir => ir.HasIncompatibleType))
                {
                    _misMatchResults.Add(caseClause);
                }
                else if (inspectedRanges.All(ir => ir.IsUnreachable))
                {
                    _unreachableResults.Add(caseClause);
                }
            }
            if (cummulativeRangeFilter.FiltersAllValues && !(_selectCaseContext.caseElseClause() is null))
            {
                _caseElseResults.Add(_selectCaseContext.caseElseClause());
            }
        }

        public List<ParserRuleContext> UnreachableCases => _unreachableResults;

        public List<ParserRuleContext> MismatchTypeCases => _misMatchResults;

        public List<ParserRuleContext> UnreachableCaseElseCases => _caseElseResults;

        private static bool InspectionCanEvaluateTypeName(string typeName) => !(typeName == string.Empty || typeName == Tokens.Variant);

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

            if (InspectionCanEvaluateTypeName(_evalTypeName))
            {
                return _evalTypeName;
            }
            _evalTypeName = DeriveTypeFromCaseClauses(inspValues, selectStmt);
            return _evalTypeName;
        }

        private string DeriveTypeFromCaseClauses(IUCIValueResults inspValues, VBAParser.SelectCaseStmtContext selectStmt)
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
                        var inspRange = _inspectionRangeFactory.Create(range, inspValues);
                        var types = inspRange.ResultContexts.Select(rc => inspValues.GetTypeName(rc));
                        caseClauseTypeNames.AddRange(types);
                    }
                }
            }

            if (TryDetermineEvaluationTypeFromTypes(caseClauseTypeNames, out _evalTypeName))
            {
                return _evalTypeName;
            }
            return string.Empty;
        }

        private static bool TryDetermineEvaluationTypeFromTypes(IEnumerable<string> typeNames, out string typeName)
        {
            typeName = string.Empty;
            var typeList = typeNames.ToList();

            //If everything is declared as a Variant, we do not attempt to inspect the selectStatement
            if (typeList.All(tn => new string[] { Tokens.Variant }.Contains(tn)))
            {
                return false;
            }
            typeList.All(tn => new string[] { typeList.First() }.Contains(tn));
            //If all match, the typeName is easy...This is the only way to return "String" or "Currency".
            if (typeList.All(tn => new string[] { typeList.First() }.Contains(tn)))
            {
                typeName = typeList.First();
                return true;
            }
            //Integer numbers will be evaluated using Long
            if (typeList.All(tn => new string[] { Tokens.Long, Tokens.Integer, Tokens.Byte }.Contains(tn)))
            {
                typeName = Tokens.Long;
                return true;
            }

            //Mix of Integertypes and rational number types will be evaluated using Double
            if (typeList.All(tn => new string[] { Tokens.Long, Tokens.Integer, Tokens.Byte, Tokens.Single, Tokens.Double }.Contains(tn)))
            {
                typeName = Tokens.Double;
                return true;
            }

            return false;
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

    internal static class MathTokens
    {
        public static readonly string MULT = "*";
        public static readonly string DIV = "/";
        public static readonly string ADD = "+";
        public static readonly string SUBTRACT = "-";
        public static readonly string POW = "^";
        public static readonly string MOD = Tokens.Mod;
        public static readonly string ADDITIVE_INVERSE = "-";
    }

    internal static class CompareTokens
    {
        public static readonly string EQ = "=";
        public static readonly string NEQ = "<>";
        public static readonly string LT = "<";
        public static readonly string LTE = "<=";
        public static readonly string GT = ">";
        public static readonly string GTE = ">=";
    }

}
