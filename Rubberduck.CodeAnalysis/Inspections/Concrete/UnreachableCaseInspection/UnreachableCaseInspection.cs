using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public sealed class UnreachableCaseInspection : ParseTreeInspectionBase
    {
        private readonly IParseTreeValueVisitorFactory _parseTreeVisitorFactory;
        private readonly ISelectCaseStmtContextWrapperFactory _selectStmtFactory;
        private readonly IParseTreeValueFactory _valueFactory;
        private enum CaseInpectionResult { Unreachable, MismatchType, CaseElse };

        private static readonly Dictionary<CaseInpectionResult, string> ResultMessages = new Dictionary<CaseInpectionResult, string>()
        {
            [CaseInpectionResult.Unreachable] = InspectionResults.UnreachableCaseInspection_Unreachable,
            [CaseInpectionResult.MismatchType] = InspectionResults.UnreachableCaseInspection_TypeMismatch,
            [CaseInpectionResult.CaseElse] = InspectionResults.UnreachableCaseInspection_CaseElse
        };

        public UnreachableCaseInspection(RubberduckParserState state) : base(state)
        {
            //TODO_Question: IUnreachableCaseInspectionFactoryFactory - candidate for IoCInstaller?
            var factoriesFactory = new UnreachableCaseInspectionFactoryProvider();

            _selectStmtFactory = factoriesFactory.CreateISelectStmtContextWrapperFactory();
            _valueFactory = factoriesFactory.CreateIParseTreeValueFactory();
            _parseTreeVisitorFactory = factoriesFactory.CreateIParseTreeValueVisitorFactory();
        }

        public override IInspectionListener Listener { get; } =
            new UnreachableCaseInspectionListener();

        private List<IInspectionResult> _inspectionResults = new List<IInspectionResult>();
        private ParseTreeVisitorResults _valueResults = new ParseTreeVisitorResults();

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            _inspectionResults = new List<IInspectionResult>();
            var qualifiedSelectCaseStmts = Listener.Contexts
                .Where(result => !IsIgnoringInspectionResultFor(result.ModuleName, result.Context.Start.Line));

            var parseTreeValueVisitor = _parseTreeVisitorFactory.Create(State, _valueFactory);
            _valueResults = new ParseTreeVisitorResults();
            parseTreeValueVisitor.OnValueResultCreated += _valueResults.OnNewValueResult;

            foreach (var qualifiedSelectCaseStmt in qualifiedSelectCaseStmts)
            {
                qualifiedSelectCaseStmt.Context.Accept(parseTreeValueVisitor);
                var selectStmt = _selectStmtFactory.Create((VBAParser.SelectCaseStmtContext)qualifiedSelectCaseStmt.Context, _valueResults);

                selectStmt.InspectForUnreachableCases();

                selectStmt.UnreachableCases.ForEach(uc => CreateInspectionResult(qualifiedSelectCaseStmt, uc, ResultMessages[CaseInpectionResult.Unreachable]));
                selectStmt.MismatchTypeCases.ForEach(mm => CreateInspectionResult(qualifiedSelectCaseStmt, mm, ResultMessages[CaseInpectionResult.MismatchType]));
                selectStmt.UnreachableCaseElseCases.ForEach(ce => CreateInspectionResult(qualifiedSelectCaseStmt, ce, ResultMessages[CaseInpectionResult.CaseElse]));
            }
            return _inspectionResults;
        }

        private void CreateInspectionResult(QualifiedContext<ParserRuleContext> selectStmt, ParserRuleContext unreachableBlock, string message)
        {
            var result = new QualifiedContextInspectionResult(this,
                        message,
                        new QualifiedContext<ParserRuleContext>(selectStmt.ModuleName, unreachableBlock));
            _inspectionResults.Add(result);
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