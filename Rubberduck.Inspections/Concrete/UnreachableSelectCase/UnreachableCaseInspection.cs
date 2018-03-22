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

namespace Rubberduck.Inspections.Concrete.UnreachableSelectCase
{
    public interface IUnreachableCaseInspectionFactoryFactory
    {
        IUnreachableCaseInspectionVisitorFactory CreateVisitorFactory();
        IUnreachableCaseInspectionValueFactory CreateValueFactory();
        IUnreachableCaseInspectionSummaryClauseFactory CreateSummaryClauseFactory();
    }

    public class UnreachableCaseInspectionFactoryFactory : IUnreachableCaseInspectionFactoryFactory
    {
        public IUnreachableCaseInspectionVisitorFactory CreateVisitorFactory()
        {
            return new UnreachableCaseInspectionVisitorFactory();
        }

        public IUnreachableCaseInspectionValueFactory CreateValueFactory()
        {
            return new UnreachableCaseInspectionValueFactory();
        }

        public IUnreachableCaseInspectionSummaryClauseFactory CreateSummaryClauseFactory()
        {
            return new UnreachableCaseInspectionSummaryClauseFactory();
        }
    }

    public class UnreachableCaseInspectionFactoryFactory2 : IUnreachableCaseInspectionFactoryFactory
    {
        public IUnreachableCaseInspectionVisitorFactory CreateVisitorFactory()
        {
            return new UnreachableCaseInspectionVisitorFactory();
        }

        public IUnreachableCaseInspectionValueFactory CreateValueFactory()
        {
            return new UnreachableCaseInspectionValueFactory();
        }

        public IUnreachableCaseInspectionSummaryClauseFactory CreateSummaryClauseFactory()
        {
            return new UnreachableCaseInspectionSummaryClauseFactory2();
        }
    }

    public sealed class UnreachableCaseInspection : ParseTreeInspectionBase
    {
        private IUnreachableCaseInspectionFactoryFactory _factoriesFactory;
        private IUnreachableCaseInspectionSummaryClauseFactory _summaryCoverageFactory;
        private IUnreachableCaseInspectionVisitorFactory _visitorFactory;
        private IUnreachableCaseInspectionValueFactory _valueFactory;
        private enum ClauseEvaluationResult { Unreachable, MismatchType, CaseElse };

        private Dictionary<ClauseEvaluationResult, string> ResultMessages = new Dictionary<ClauseEvaluationResult, string>()
        {
            [ClauseEvaluationResult.Unreachable] = InspectionsUI.UnreachableCaseInspection_Unreachable,
            [ClauseEvaluationResult.MismatchType] = InspectionsUI.UnreachableCaseInspection_TypeMismatch,
            [ClauseEvaluationResult.CaseElse] = InspectionsUI.UnreachableCaseInspection_CaseElse
        };

        public UnreachableCaseInspection(RubberduckParserState state) : base(state)
        {
            //TODO_Question: Candidate(s) for IoCInstaller?  Or...not appropriate?
            _summaryCoverageFactory = FactoriesFactory.CreateSummaryClauseFactory();
            _visitorFactory = FactoriesFactory.CreateVisitorFactory();
            _valueFactory = FactoriesFactory.CreateValueFactory();
        }

        public IUnreachableCaseInspectionFactoryFactory FactoriesFactory
        {
            get
            {
                if(_factoriesFactory is null)
                {
                    _factoriesFactory = new UnreachableCaseInspectionFactoryFactory();
                }
                return _factoriesFactory;
            }
            set
            {
                if(value != _factoriesFactory)
                {
                    _factoriesFactory = value;
                    _summaryCoverageFactory = _factoriesFactory.CreateSummaryClauseFactory();
                    _visitorFactory = _factoriesFactory.CreateVisitorFactory();
                    _valueFactory = _factoriesFactory.CreateValueFactory();
                }
            }
        }


        public override IInspectionListener Listener { get; } =
            new UnreachableCaseInspectionListener();

        private List<IInspectionResult> InspectionResults { set; get; } = new List<IInspectionResult>();

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var qualifiedSelectCaseStmts = Listener.Contexts
                .Where(result => !IsIgnoringInspectionResultFor(result.ModuleName, result.Context.Start.Line));

            foreach (var qualifiedSelectCaseStmt in qualifiedSelectCaseStmts)
            {
                var selectCaseVisitor = _visitorFactory.Create(State);
                var selectCaseContext = (VBAParser.SelectCaseStmtContext)qualifiedSelectCaseStmt.Context;
                var selectStmtValueResults = selectCaseContext.Accept(selectCaseVisitor);
                IUnreachableCaseInspectionSelectStmt selectStmt = new UnreachableCaseInspectionSelectStmt(selectCaseContext, selectStmtValueResults, _summaryCoverageFactory, _valueFactory);

                selectStmt.EvaluateForUnreachableCases();
                selectStmt.UnreachableCases.ForEach(uc => CreateInspectionResult(qualifiedSelectCaseStmt, uc, ResultMessages[ClauseEvaluationResult.Unreachable]));
                selectStmt.MismatchTypeCases.ForEach(mm => CreateInspectionResult(qualifiedSelectCaseStmt, mm, ResultMessages[ClauseEvaluationResult.MismatchType]));
                selectStmt.UnreachableCaseElseCases.ForEach(ce => CreateInspectionResult(qualifiedSelectCaseStmt, ce, ResultMessages[ClauseEvaluationResult.CaseElse]));
            }
            return InspectionResults;
        }

        private void CreateInspectionResult(QualifiedContext<ParserRuleContext> selectStmt, ParserRuleContext unreachableBlock, string message)
        {
            var result = new QualifiedContextInspectionResult(this,
                        message,
                        new QualifiedContext<ParserRuleContext>(selectStmt.ModuleName, unreachableBlock));
            InspectionResults.Add(result);
        }

        #region UnreachableCaseInspectionListeners
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