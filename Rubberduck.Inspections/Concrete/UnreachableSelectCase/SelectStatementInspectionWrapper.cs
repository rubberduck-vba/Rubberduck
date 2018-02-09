using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Inspections.Concrete
{
    public class SelectStatementInspectionWrapper
    {
        private readonly RubberduckParserState _state;
        private readonly QualifiedContext<ParserRuleContext> _qSelectStmt;
        private readonly VBAParser.SelectCaseStmtContext _selectStmt;

        private string _evaluationTypeName;
        private SelectExpressionContextVisitor _selectExprContextVisitor;

        public SelectStatementInspectionWrapper(QualifiedContext<ParserRuleContext> selectStmt, RubberduckParserState state)
        {
            _state = state;
            _qSelectStmt = selectStmt;
            _selectStmt = selectStmt.Context as VBAParser.SelectCaseStmtContext;
            _evaluationTypeName = string.Empty;
            _selectExprContextVisitor = new SelectExpressionContextVisitor(state);
            _evaluationTypeName = GetSelectCaseEvaluationType(_selectStmt, _selectExprContextVisitor);
        }

        public string EvaluationTypeName => _evaluationTypeName;
        public bool CanBeInspected => !EvaluationTypeName.Equals(string.Empty);
        public QualifiedContext<ParserRuleContext> QualifiedContext => _qSelectStmt;
        public IEnumerable<VBAParser.CaseClauseContext> CaseClauses => _selectStmt.caseClause();

        public bool HasCaseElse => !(CaseElse is null);
        public VBAParser.CaseElseClauseContext CaseElse => _selectStmt.caseElseClause();

        private static string GetSelectCaseEvaluationType(VBAParser.SelectCaseStmtContext selectStmt, SelectExpressionContextVisitor selectStmtTypeVisitor)
        {
            return selectStmt.Accept(selectStmtTypeVisitor);
        }
    }
}
