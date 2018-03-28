using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public sealed class UnreachableCaseInspection : ParseTreeInspectionBase
    {
        private IUCIParseTreeValueVisitorFactory _parseTreeVisitorFactory;
        private IUnreachableCaseInspectionSelectStmtFactory _selectStmtFactory;
        private IUCIValueFactory _valueFactory;
        private IUCIParseTreeValueVisitor _parseTreeValueVisitor;
        private enum CaseInpectionResult { Unreachable, MismatchType, CaseElse };

        private static Dictionary<CaseInpectionResult, string> ResultMessages = new Dictionary<CaseInpectionResult, string>()
        {
            [CaseInpectionResult.Unreachable] = InspectionsUI.UnreachableCaseInspection_Unreachable,
            [CaseInpectionResult.MismatchType] = InspectionsUI.UnreachableCaseInspection_TypeMismatch,
            [CaseInpectionResult.CaseElse] = InspectionsUI.UnreachableCaseInspection_CaseElse
        };

        public UnreachableCaseInspection(RubberduckParserState state) : base(state)
        {
            //TODO_Question: IUnreachableCaseInspectionFactoryFactory - candidate for IoCInstaller?
            var factoriesFactory = new UnreachableCaseInspectionFactoryFactory();

            _selectStmtFactory = factoriesFactory.CreateUnreachableCaseInspectionSelectStmtFactory();
            _valueFactory = factoriesFactory.CreateIUCIValueFactory();
            _parseTreeVisitorFactory = factoriesFactory.CreateIUCIParseTreeValueVisitorFactory();

            _parseTreeValueVisitor = _parseTreeVisitorFactory.Create(State, _valueFactory);
        }

        public override IInspectionListener Listener { get; } =
            new UnreachableCaseInspectionListener();

        private List<IInspectionResult> InspectionResults { set; get; } = new List<IInspectionResult>();

        private UCIValueResults ValueResults { get; } = new UCIValueResults();

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var qualifiedSelectCaseStmts = Listener.Contexts
                .Where(result => !IsIgnoringInspectionResultFor(result.ModuleName, result.Context.Start.Line));

            var parseTreeValueVisitor = _parseTreeVisitorFactory.Create(State, _valueFactory);
            parseTreeValueVisitor.OnValueResultCreated += ValueResults.OnNewValueResult;

            foreach (var qualifiedSelectCaseStmt in qualifiedSelectCaseStmts)
            {
                qualifiedSelectCaseStmt.Context.Accept(parseTreeValueVisitor);
                var selectStmt = _selectStmtFactory.Create((VBAParser.SelectCaseStmtContext)qualifiedSelectCaseStmt.Context, ValueResults);

                selectStmt.InspectForUnreachableCases();

                selectStmt.UnreachableCases.ForEach(uc => CreateInspectionResult(qualifiedSelectCaseStmt, uc, ResultMessages[CaseInpectionResult.Unreachable]));
                selectStmt.MismatchTypeCases.ForEach(mm => CreateInspectionResult(qualifiedSelectCaseStmt, mm, ResultMessages[CaseInpectionResult.MismatchType]));
                selectStmt.UnreachableCaseElseCases.ForEach(ce => CreateInspectionResult(qualifiedSelectCaseStmt, ce, ResultMessages[CaseInpectionResult.CaseElse]));
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