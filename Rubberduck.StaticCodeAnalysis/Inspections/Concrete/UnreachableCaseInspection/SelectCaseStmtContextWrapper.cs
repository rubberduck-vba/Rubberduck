using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface ISelectCaseStmtContextWrapper
    {
        string EvaluationTypeName { get; }
        void InspectForUnreachableCases();
        List<ParserRuleContext> UnreachableCases { get; }
        List<ParserRuleContext> MismatchTypeCases { get; }
        List<ParserRuleContext> UnreachableCaseElseCases { get; }
    }

    public class SelectCaseStmtContextWrapper : ContextWrapperBase, ISelectCaseStmtContextWrapper
    {
        private readonly VBAParser.SelectCaseStmtContext _selectCaseContext;
        private IRangeClauseContextWrapperFactory _inspectionRangeFactory;
        private List<ParserRuleContext> _unreachableResults;
        private List<ParserRuleContext> _mismatchResults;
        private List<ParserRuleContext> _caseElseResults;

        public SelectCaseStmtContextWrapper(VBAParser.SelectCaseStmtContext selectCaseContext, IParseTreeVisitorResults inspValues, IUnreachableCaseInspectionFactoryProvider factoryFactory)
            : base(selectCaseContext, inspValues, factoryFactory)
        {
            _selectCaseContext = selectCaseContext;
            _unreachableResults = new List<ParserRuleContext>();
            _mismatchResults = new List<ParserRuleContext>();
            _caseElseResults = new List<ParserRuleContext>();
            _inspectionRangeFactory = factoryFactory.CreateIRangeClauseContextWrapperFactory();
            SetEvaluationTypeName(Context, ParseTreeValueResults, FilterFactory);
        }

        public string EvaluationTypeName { private set; get; }

        public void InspectForUnreachableCases()
        {
            if (!InspectionCanEvaluateTypeName(EvaluationTypeName))
            {
                return;
            }

            var cummulativeRangeFilter = FilterFactory.Create(EvaluationTypeName, ValueFactory);
            foreach (var caseClause in _selectCaseContext.caseClause())
            {
                var inspectedRanges = new List<IRangeClauseContextWrapper>();
                foreach (var range in caseClause.rangeClause())
                {
                    var inspectionRange = _inspectionRangeFactory.Create(range, EvaluationTypeName, _inspValues);
                    if (inspectionRange.IsReachable(cummulativeRangeFilter))
                    {
                        cummulativeRangeFilter.Add(inspectionRange.AsFilter);
                    }
                    inspectedRanges.Add(inspectionRange);
                }

                if (inspectedRanges.All(ir => ir.HasIncompatibleType))
                {
                    _mismatchResults.Add(caseClause);
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

        public List<ParserRuleContext> MismatchTypeCases => _mismatchResults;

        public List<ParserRuleContext> UnreachableCaseElseCases => _caseElseResults;

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
            Tokens.String
        };

        private static bool InspectionCanEvaluateTypeName(string typeName) => InspectableTypes.Contains(typeName);

        private void SetEvaluationTypeName(ParserRuleContext context, IParseTreeVisitorResults inspValues, IRangeClauseFilterFactory factory)
        {
            var selectStmt = (VBAParser.SelectCaseStmtContext)context;
            if (TryDetectTypeHint(selectStmt.selectExpression().GetText(), out string evalTypName))
            {
                EvaluationTypeName = evalTypName;
                return;
            }

            var typeName = string.Empty;
            if (inspValues.TryGetValue(selectStmt.selectExpression(), out IParseTreeValue result))
            {
                EvaluationTypeName = result.TypeName;
            }

            if (InspectionCanEvaluateTypeName(EvaluationTypeName))
            {
                return;
            }
            EvaluationTypeName = DeriveTypeFromCaseClauses(inspValues, selectStmt);
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
                        var inspRange = _inspectionRangeFactory.Create(range, inspValues);
                        var types = inspRange.ResultContexts.Select(rc => inspValues.GetTypeName(rc))
                            .Where(tp => InspectableTypes.Contains(tp));
                        caseClauseTypeNames.AddRange(types);
                    }
                }
            }

            if (TryDetermineEvaluationTypeFromTypes(caseClauseTypeNames, out string evalTypeName))
            {
                return evalTypeName;
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

    internal static class LogicSymbols
    {
        private static string _lessThan;
        private static string _greaterThan;
        private static string _equalTo;

        public static string EQ => _equalTo ?? LoadSymbols(VBAParser.EQ);
        public static string NEQ => "<>";
        public static string LT => _lessThan ?? LoadSymbols(VBAParser.LT);
        public static string LTE => "<=";
        public static string GT => _greaterThan ?? LoadSymbols(VBAParser.GT);
        public static string GTE => ">=";
        public static string AND => Tokens.And;
        public static string OR => Tokens.Or;
        public static string XOR => Tokens.XOr;
        public static string NOT => Tokens.Not;
        public static string EQV => "Eqv";
        public static string IMP => "Imp";

        private static string LoadSymbols(int target)
        {
            _lessThan = VBAParser.DefaultVocabulary.GetLiteralName(VBAParser.LT).Replace("'", "");
            _greaterThan = VBAParser.DefaultVocabulary.GetLiteralName(VBAParser.GT).Replace("'", "");
            _equalTo = VBAParser.DefaultVocabulary.GetLiteralName(VBAParser.EQ).Replace("'", "");
            return VBAParser.DefaultVocabulary.GetLiteralName(target).Replace("'", "");
        }
    }
}
