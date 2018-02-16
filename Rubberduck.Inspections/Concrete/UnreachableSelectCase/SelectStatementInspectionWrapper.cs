using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Inspections.Concrete
{
    //public class SelectStatementInspectionWrapper
    //{
    //    private readonly QualifiedContext<ParserRuleContext> _qSelectStmt;
    //    private readonly VBAParser.SelectCaseStmtContext _selectStmt;
    //    private string _evaluationType;

    //    private SelectExpressionContextVisitor _selectExprContextVisitor;

    //    public SelectStatementInspectionWrapper(QualifiedContext<ParserRuleContext> selectStmt, ContextValueVisitor ctxtValueVisitor, IParseTreeValueResults parseResults)
    //    {
    //        _qSelectStmt = selectStmt;
    //        _selectStmt = selectStmt.Context as VBAParser.SelectCaseStmtContext;
    //        _selectExprContextVisitor = new SelectExpressionContextVisitor(parseResults);
    //        _evaluationType = GetSelectCaseEvaluationType(_selectStmt, _selectExprContextVisitor);
    //    }

    //    public string EvaluationTypeName => _evaluationType;
    //    public bool CanBeInspected => !EvaluationTypeName.Equals(string.Empty);
    //    public QualifiedContext<ParserRuleContext> QualifiedContext => _qSelectStmt;
    //    public IEnumerable<VBAParser.CaseClauseContext> CaseClauses => _selectStmt.caseClause();

    //    public bool HasCaseElse => !(CaseElse is null);
    //    public VBAParser.CaseElseClauseContext CaseElse => _selectStmt.caseElseClause();

    //    private static string GetSelectCaseEvaluationType(VBAParser.SelectCaseStmtContext selectStmt, SelectExpressionContextVisitor selectStmtTypeVisitor)
    //    {
    //        return selectStmt.Accept(selectStmtTypeVisitor);
    //    }
    //}
}
