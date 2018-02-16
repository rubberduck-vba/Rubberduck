using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime.Tree;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace Rubberduck.Inspections.Concrete
{
    public sealed class UnreachableCaseInspection : ParseTreeInspectionBase
    {
        private enum ClauseEvaluationResult { Unreachable, MismatchType, CaseElse };

        private Dictionary<ClauseEvaluationResult, string> ResultMessages = new Dictionary<ClauseEvaluationResult, string>()
        {
            [ClauseEvaluationResult.Unreachable] = InspectionsUI.UnreachableCaseInspection_Unreachable,
            [ClauseEvaluationResult.MismatchType] = InspectionsUI.UnreachableCaseInspection_TypeMismatch,
            [ClauseEvaluationResult.CaseElse] = InspectionsUI.UnreachableCaseInspection_CaseElse
        };

        private struct SelectCaseData
        {
            public Dictionary<QualifiedContext<ParserRuleContext>,string> SelectCaseAndTypeNames;
            public IParseTreeValueResults ParseTreeResults;
        }

        public UnreachableCaseInspection(RubberduckParserState state) : base(state) { }

        public override IInspectionListener Listener { get; } =
            new UnreachableCaseInspectionListener();

        private List<IInspectionResult> InspectionResults { set; get; } = new List<IInspectionResult>();

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            InspectionResults = new List<IInspectionResult>();
            var selectContextData = LoadSelectCaseData();
            foreach (var selectCaseCtxtTypeName in selectContextData.SelectCaseAndTypeNames.Where(sc => sc.Value != string.Empty))
            {
                var selectCaseSummarizedCoverage = UnreachableSelectCaseFactory.CreateSummaryCoverage(
                    selectCaseCtxtTypeName.Key.Context, 
                    selectContextData.ParseTreeResults, 
                    selectCaseCtxtTypeName.Value);

                InspectSelectStatement(selectCaseCtxtTypeName.Key, selectCaseSummarizedCoverage);
            }

            return InspectionResults;
        }

        private SelectCaseData LoadSelectCaseData()
        {
            var selectCaseStmtContexts = Listener.Contexts
                .Where(result => !IsIgnoringInspectionResultFor(result.ModuleName, result.Context.Start.Line));

            var selectData = new SelectCaseData()
            {
                SelectCaseAndTypeNames = new Dictionary<QualifiedContext<ParserRuleContext>, string>()
            };

            foreach (var selectCaseStmt in selectCaseStmtContexts)
            {
                //TODO: Maybe - ISelectCaseTypeEvaluator (returns typeName)
                var selectCaseUnTypedVisitor = UnreachableSelectCaseFactory.CreateParseTreeVisitor(State);
                var selectCaseResults = selectCaseStmt.Context.Accept(selectCaseUnTypedVisitor);
                var evaluationTypeName = DetermineSelectCaseEvaluationTypeName(selectCaseStmt.Context, selectCaseResults);

                var selectCaseTypedVisitor = UnreachableSelectCaseFactory.CreateParseTreeVisitor(State, evaluationTypeName);
                var parseTreeValueResults = selectCaseStmt.Context.Accept(selectCaseTypedVisitor);

                selectData.SelectCaseAndTypeNames.Add(selectCaseStmt, evaluationTypeName);
                if (selectData.ParseTreeResults is null)
                {
                    selectData.ParseTreeResults = parseTreeValueResults;
                }
                else
                {
                    selectData.ParseTreeResults.Add(parseTreeValueResults);
                }
            }
            return selectData;
        }

        private void InspectSelectStatement(QualifiedContext<ParserRuleContext> qualifiedSelectCaseContext, ISummaryCoverage allSelectCaseContextCoverage)
        {
            //var inspResults = new List<IInspectionResult>();
            if(!(qualifiedSelectCaseContext.Context is VBAParser.SelectCaseStmtContext selectCaseStatementContext))
            {
                throw new ArgumentException("Invalid argument type", "selectCaseStatementContext");
            }

            var cummulativeCoverage = UnreachableSelectCaseFactory.CreateSummaryCoverageShell(allSelectCaseContextCoverage.TypeName);
            selectCaseStatementContext = qualifiedSelectCaseContext.Context as VBAParser.SelectCaseStmtContext;
            foreach (var caseClauseCtxt in selectCaseStatementContext.caseClause())
            {
                if (cummulativeCoverage.CoversAllValues)
                {
                    //Once all values are covered, the remaining CaseClauses are unreachable
                    //and there is no point in evaluating them
                    CreateInspectionResult(qualifiedSelectCaseContext, caseClauseCtxt, ResultMessages[ClauseEvaluationResult.Unreachable]);
                    //inspResults.Add(result);
                    continue;
                }

                if (allSelectCaseContextCoverage.CanBeInspected(caseClauseCtxt.rangeClause()))
                {
                    var caseClauseCoverge = allSelectCaseContextCoverage.CoverageFor(caseClauseCtxt);
                    if (caseClauseCoverge.HasConditionsNotCoveredBy(cummulativeCoverage, out ISummaryCoverage additionalCoverage))
                    {
                        cummulativeCoverage.Add(additionalCoverage);
                    }
                    else
                    {
                        //If there are no conditions to add, then the current CaseClause's conditions are covered
                        //by the combination of prior CaseClauses - and the CaseClause is therefore unreachable
                        CreateInspectionResult(qualifiedSelectCaseContext, caseClauseCtxt, ResultMessages[ClauseEvaluationResult.Unreachable]);
                        //inspResults.Add(result);
                    }
                }
                else
                {
                    //Call out CaseClauses that cannot be implicitly converted to the SelectCase type as a special case of unreachable
                    if (allSelectCaseContextCoverage.IsIncompatibleType(caseClauseCtxt.rangeClause()))
                    {
                        CreateInspectionResult(qualifiedSelectCaseContext, caseClauseCtxt, ResultMessages[ClauseEvaluationResult.MismatchType]);
                        //inspResults.Add(result);
                    }
                }
            }

            if (cummulativeCoverage.CoversAllValues && !(selectCaseStatementContext.caseElseClause() is null))
            {
                CreateInspectionResult(qualifiedSelectCaseContext, selectCaseStatementContext.caseElseClause(), ResultMessages[ClauseEvaluationResult.CaseElse]);
                //inspResults.Add(result);
            }
            //return inspResults;
        }

        private void CreateInspectionResult(QualifiedContext<ParserRuleContext> selectStmt, ParserRuleContext unreachableBlock, string message)
        {
            var result = new QualifiedContextInspectionResult(this,
                        message,
                        new QualifiedContext<ParserRuleContext>(selectStmt.ModuleName, unreachableBlock));
            InspectionResults.Add(result);
        }

        public static string DetermineSelectCaseEvaluationTypeName(ParserRuleContext selectStmt, IParseTreeValueResults selectStmtValues)
        {
            Debug.Assert(selectStmt is VBAParser.SelectCaseStmtContext);

            var selectExpression = ((VBAParser.SelectCaseStmtContext)selectStmt).selectExpression();
            if (selectExpression.children.Any(child => IsLogicalContext(child) || IsTrueFalseLiteral(child)))
            {
                return Tokens.Boolean;
            }

            if (selectExpression.children.Any(child => child is VBAParser.ConcatOpContext))
            {
                return Tokens.String;
            }

            var theTypeName = string.Empty;
            var smplName = selectExpression.GetDescendent<VBAParser.SimpleNameExprContext>();
            if (SymbolList.TypeHintToTypeName.TryGetValue(smplName.GetText().Last().ToString(), out theTypeName))
            {
                return theTypeName;
            }

            var selectExpressionContexts = selectStmtValues.AllContexts.Where(se => se.IsDescendentOf<VBAParser.SelectExpressionContext>());
            if (selectStmtValues.AllContexts.Any(se => selectStmtValues.Result(se).HasValue))
            {
                var unresolvedContextTypeNames = selectExpressionContexts.Where(val => selectStmtValues.Result(val).HasDeclaredTypeName).Select(val => selectStmtValues.Result(val).DeclaredTypeName);
                if (TryDetermineEvaluationTypeFromTypes(unresolvedContextTypeNames, out theTypeName))
                {
                    return theTypeName;
                }
            }
            else
            {
                var resolvedContextTypeNames = selectExpressionContexts.Where(val => selectStmtValues.Result(val).HasDeclaredTypeName).Select(val => selectStmtValues.Result(val).DeclaredTypeName);

                if (TryDetermineEvaluationTypeFromTypes(resolvedContextTypeNames, out theTypeName))
                {
                    return theTypeName;
                }
            }

            var typeNames = selectStmtValues.RangeClauseResults().Select(res => res.UseageTypeName);
            if (TryDetermineEvaluationTypeFromTypes(typeNames, out string typeName))
            {
                return typeName;
            }

            //If Strings are in the mix and prevent resolution to a type, we remove them
            //here and see if a resolution becomes possible.  The strings will be converted to the
            //final type during subsequent unreachable analysis.  If they cannot be converted to
            //the "Evaluation Type", they will be flagged as mismatching e.g., "45" converts to a number
            //but "foo" will not.
            var modifiedNames = typeNames.ToList();
            modifiedNames.RemoveAll(tn => tn.Equals(Tokens.String));
            if (TryDetermineEvaluationTypeFromTypes(modifiedNames, out typeName))
            {
                return typeName;
            }
            return string.Empty;
        }

        private static bool TryDetermineEvaluationTypeFromTypes(IEnumerable<string> typeNames, out string typeName)
        {
            typeName = string.Empty;
            var typeList = typeNames.ToList();
            typeList.Remove(Tokens.Variant);
            if (!typeList.Any())
            {
                return false;
            }
            //To select "String" or "Currency", all types in the typelist must match
            if (typeList.All(tn => tn.Equals(typeList.First())))
            {
                typeName = typeList.First();
                return true;
            }

            var nextType = new string[] { Tokens.Long, Tokens.LongLong, Tokens.Integer, Tokens.Byte };
            var result = typeList.All(tn => nextType.Contains(tn));
            if (result)
            {
                typeName = Tokens.Long;
                return true;
            }

            nextType = new string[] { Tokens.Long, Tokens.LongLong, Tokens.Integer, Tokens.Byte, Tokens.Single, Tokens.Double };
            result = typeList.All(tn => nextType.Contains(tn));
            if (result)
            {
                typeName = Tokens.Double;
                return true;
            }
            return false;
        }

        private static bool IsLogicalContext<T>(T child)
        {
            return child is VBAParser.RelationalOpContext
                || child is VBAParser.LogicalXorOpContext
                || child is VBAParser.LogicalAndOpContext
                || child is VBAParser.LogicalOrOpContext
                || child is VBAParser.LogicalEqvOpContext
                || child is VBAParser.LogicalNotOpContext;
        }

        private static bool IsTrueFalseLiteral<T>(T child)
        {
            if (child is VBAParser.LiteralExprContext)
            {
                var litExpr = child as VBAParser.LiteralExprContext;
                return litExpr.GetText().Equals(Tokens.True) || litExpr.GetText().Equals(Tokens.False);
            }
            return false;
        }


        #region UnreachableCaseInspectionListener
        public class UnreachableCaseInspectionListener : VBAParserBaseListener, IInspectionListener
        {
            private readonly List<QualifiedContext<ParserRuleContext>> _contexts = new List<QualifiedContext<ParserRuleContext>>();
            public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;

            public QualifiedModuleName CurrentModuleName { get; set; }

            public void ClearContexts()
            {
                _contexts.Clear();
            }

            public override void EnterSelectCaseStmt([NotNull] VBAParser.SelectCaseStmtContext context)
            {
                _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
            }
        }
        #endregion
    }
}