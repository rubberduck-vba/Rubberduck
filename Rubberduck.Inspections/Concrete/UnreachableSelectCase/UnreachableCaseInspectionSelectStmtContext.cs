using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Inspections.Concrete.UnreachableSelectCase
{
    public interface IUnreachableCaseInspectionContext
    {
        ParserRuleContext Context { get; }
    }

    public interface IUnreachableCaseInspectionSelectStmt : IUnreachableCaseInspectionContext
    {
        string EvaluationTypeName { set; get; }
        QualifiedContext<ParserRuleContext> QualifiedContext { get; }
        bool CanBeInspected { get; }
        List<IUnreachableCaseInspectionCaseClause> CaseClauses { get; }
    }

    public class UnreachableCaseInspectionSelectStmtContext : UnreachableCaseInspectionContext, IUnreachableCaseInspectionSelectStmt
    {
        private readonly QualifiedContext<ParserRuleContext> _qualifiedContext;
        private readonly IParseTreeVisitor<IUnreachableCaseInspectionValue> _ptVisitor;
        private string _evaluationTypeName;
        private List<IUnreachableCaseInspectionCaseClause> _caseClauses;
        public UnreachableCaseInspectionSelectStmtContext(QualifiedContext<ParserRuleContext> qualifiedContext, IParseTreeVisitor<IUnreachableCaseInspectionValue> ptVisitor) : base(qualifiedContext.Context)
        {
            _qualifiedContext = qualifiedContext;
            _ptVisitor = ptVisitor;
            _evaluationTypeName = null;
            _caseClauses = new List<IUnreachableCaseInspectionCaseClause>();
            var selectStmt = (VBAParser.SelectCaseStmtContext)qualifiedContext.Context;
            foreach(var caseClause in selectStmt.caseClause())
            {
                _caseClauses.Add(new UnreachableCaseInspectionCaseClause(caseClause));
            }
        }

        public List<IUnreachableCaseInspectionCaseClause> CaseClauses => _caseClauses;
        public bool CanBeInspected => !(EvaluationTypeName.Equals(string.Empty) || EvaluationTypeName.Equals(Tokens.Variant));

        public string EvaluationTypeName
        {
            set
            {
                _evaluationTypeName = value;
            }

            get
            {
                if(_evaluationTypeName is null)
                {
                    _evaluationTypeName = DetermineSelectCaseEvaluationTypeName((VBAParser.SelectCaseStmtContext)Context, _ptVisitor);
                }
                return _evaluationTypeName;
            }
        }
        public QualifiedContext<ParserRuleContext> QualifiedContext => _qualifiedContext;

        private static string DetermineSelectCaseEvaluationTypeName(VBAParser.SelectCaseStmtContext selectStmt,  IParseTreeVisitor<IUnreachableCaseInspectionValue> ptVisitor)
        {
            var selectExpression = selectStmt.selectExpression();

            var selectExprTypeVisitor = new SelectCaseContextTypeVisitor<VBAParser.SelectExpressionContext>(ptVisitor);
            var typeName = selectExpression.Accept(selectExprTypeVisitor);
            if (typeName == string.Empty || typeName == Tokens.Variant)
            {
                var caseClauseTypeNames = new List<string>();
                foreach (var caseClause in selectStmt.caseClause())
                {
                    var caseClauseTypeVisitor = new CaseClauseTypeVisitor(caseClause, ptVisitor);
                    var caseClauseType = caseClause.Accept(caseClauseTypeVisitor);
                    caseClauseTypeNames.Add(caseClauseType);
                }

                if (TryDetermineEvaluationTypeFromTypes(caseClauseTypeNames, out typeName))
                {
                    return typeName;
                }
            }
            else
            {
                return typeName;
            }

            return string.Empty;
        }

        private static bool TryDetermineEvaluationTypeFromTypes(IEnumerable<string> typeNames, out string typeName)
        {
            typeName = string.Empty;
            var typeList = typeNames.ToList();

            //If everything is declared as a Variant, we do not attempt to inspect the selectStatement
            if (CheckAllTypesAreContainedIn(typeList, new string[] { Tokens.Variant }))
            {
                return false;
            }

            //If all match, the typeName is easy...This is the only way to return "String" or "Currency".
            if (CheckAllTypesAreContainedIn(typeList, new string[] { typeList.First() }))
            {
                typeName = typeList.First();
                return true;
            }
            //Integer numbers will be evaluated using Long
            if (CheckAllTypesAreContainedIn(typeList, new string[] { Tokens.Long, Tokens.Integer, Tokens.Byte }))
            {
                typeName = Tokens.Long;
                return true;
            }

            //Mix of Integertypes and rational number types will be evaluated using Double
            if (CheckAllTypesAreContainedIn(typeList, new string[] { Tokens.Long, Tokens.Integer, Tokens.Byte, Tokens.Single, Tokens.Double }))
            {
                typeName = Tokens.Double;
                return true;
            }

            return false;
        }

        private static bool CheckAllTypesAreContainedIn(List<string> typeList, string[] typesToUse)
        {
            return typeList.All(tn => typesToUse.Contains(tn));
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
            if (child is VBAParser.LiteralExprContext litExpr)
            {
                return litExpr.GetText().Equals(Tokens.True) || litExpr.GetText().Equals(Tokens.False);
            }
            return false;
        }
    }

    public class UnreachableCaseInspectionContext : IParseTree, IUnreachableCaseInspectionContext
    {
        protected readonly ParserRuleContext _context;
        public UnreachableCaseInspectionContext(ParserRuleContext context)
        {
            _context = context;
        }

        public TContext GetChild<TContext>() where TContext : ParserRuleContext
        {
            return Context.GetChild<TContext>();
        }

        public ParserRuleContext Context => _context;

        public IParseTree Parent => ((IParseTree)_context).Parent;

        public Interval SourceInterval => ((IParseTree)_context).SourceInterval;

        public object Payload => ((IParseTree)_context).Payload;

        public int ChildCount => ((IParseTree)_context).ChildCount;

        ITree ITree.Parent => ((IParseTree)_context).Parent;

        public virtual T Accept<T>(IParseTreeVisitor<T> visitor)
        {
            return ((IParseTree)_context).Accept(visitor);
        }

        public IParseTree GetChild(int i)
        {
            return ((IParseTree)_context).GetChild(i);
        }

        public string GetText()
        {
            return _context.GetText();
        }

        public string ToStringTree(Parser parser)
        {
            return _context.ToStringTree(parser);
        }

        public string ToStringTree()
        {
            return _context.ToStringTree();
        }

        ITree ITree.GetChild(int i)
        {
            return ((IParseTree)_context).GetChild(i);
        }
    }
}
