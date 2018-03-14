using Antlr4.Runtime;
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
    internal class CaseClauseTypeVisitor : IParseTreeVisitor<string>
    {
        private readonly SelectCaseContextTypeVisitor<VBAParser.RangeClauseContext> _rangeTypeVisitor;
        private readonly IParseTreeVisitor<IUnreachableCaseInspectionValue> _ptVisitor;
        private readonly VBAParser.CaseClauseContext _caseClause;
        public CaseClauseTypeVisitor(VBAParser.CaseClauseContext caseClauseContext, IParseTreeVisitor<IUnreachableCaseInspectionValue> ptVisitor)
        {
            _caseClause = caseClauseContext;
            _ptVisitor = ptVisitor;
            _rangeTypeVisitor = new SelectCaseContextTypeVisitor<VBAParser.RangeClauseContext>(ptVisitor);
        }

        public string Visit(IParseTree tree)
        {
            return _rangeTypeVisitor.Visit(tree);
        }

        public string VisitChildren(IRuleNode node)
        {
            var rangeClauses = _caseClause.rangeClause();
            var rangeClauseTypes = new List<string>();
            UnreachableCaseInspectionRange wrapped = null;
            foreach( var rangeClause in rangeClauses)
            {
                wrapped = new UnreachableCaseInspectionRange(rangeClause);
                if (wrapped.IsValueRange)
                {
                    var startContext = rangeClause.GetChild<VBAParser.SelectStartValueContext>();
                    var endContext = rangeClause.GetChild<VBAParser.SelectEndValueContext>();
                    var startVisitor = new SelectCaseContextTypeVisitor<VBAParser.SelectStartValueContext>(_ptVisitor);
                    var endVisitor = new SelectCaseContextTypeVisitor<VBAParser.SelectEndValueContext>(_ptVisitor);
                    rangeClauseTypes.Add(startContext.Accept(startVisitor));
                    rangeClauseTypes.Add(endContext.Accept(endVisitor));
                }
                else
                {
                    rangeClauseTypes.Add(rangeClause.Accept(_rangeTypeVisitor));
                }
            }

            if(TryDetermineEvaluationTypeFromTypes(rangeClauseTypes, out string typeName))
            {
                return typeName;
            }
            return string.Empty;
        }

        public string VisitErrorNode(IErrorNode node)
        {
            return _rangeTypeVisitor.VisitErrorNode(node);
        }

        public string VisitTerminal(ITerminalNode node)
        {
            return _rangeTypeVisitor.VisitTerminal(node);
        }

        private static bool TryDetermineEvaluationTypeFromTypes(IEnumerable<string> typeNames, out string typeName)
        {
            typeName = string.Empty;
            var typeList = typeNames.ToList();

            //Variant SelectCase statements are not evaluated
            if (CheckAllTypesAreContainedIn(typeList, new string[] { Tokens.Variant }))
            {
                return false;
            }

            //If all match, easy to choose.  
            //Note: This is the only way to return "String" or "Currency".
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
    }

    internal class SelectCaseContextTypeVisitor<T> : IParseTreeVisitor<string> where T: ParserRuleContext
    {
        private static string DEFAULT_TYPENAME = string.Empty;
        IParseTreeVisitor<IUnreachableCaseInspectionValue> _ptVisitor;
        public SelectCaseContextTypeVisitor(IParseTreeVisitor<IUnreachableCaseInspectionValue> ptVisitor)
        {
            _ptVisitor = ptVisitor;
        }

        public string Visit(IParseTree tree)
        {
            return DEFAULT_TYPENAME;
        }

        public string VisitChildren(IRuleNode node)
        {
            if(node is T)
            {
                var result = _ptVisitor.Visit(node);
                if (result.TypeName == string.Empty || result.TypeName == Tokens.Variant)
                {
                    var theTypeName = string.Empty;
                    var smplName = node.GetDescendent<VBAParser.SimpleNameExprContext>();
                    if (SymbolList.TypeHintToTypeName.TryGetValue(smplName.GetText().Last().ToString(), out theTypeName))
                    {
                        return theTypeName;
                    }
                }
                return result.TypeName;
            }
            return DEFAULT_TYPENAME;
        }

        public string VisitErrorNode(IErrorNode node)
        {
            return DEFAULT_TYPENAME;
        }

        public string VisitTerminal(ITerminalNode node)
        {
            return DEFAULT_TYPENAME;
        }
    }

    public class CaseClauseSummaryVisitor : IParseTreeVisitor<ISummaryCoverage>
    {
        private VBAParser.CaseClauseContext _caseClauseCtxt;
        private UnreachableCaseInspectionValueVisitor _visitor;
        string _typeName;
        private SummaryCoverageFactory _summaryCoverageFactory;
        public CaseClauseSummaryVisitor(VBAParser.CaseClauseContext caseClauseContext, RubberduckParserState state, string typeName)
        {
            _caseClauseCtxt = caseClauseContext;
            _visitor = new UnreachableCaseInspectionValueVisitor(state, new IUnreachableCaseInspectionValueFactory());
            _typeName = typeName;
            _summaryCoverageFactory = new SummaryCoverageFactory();
        }

        public ISummaryCoverage Visit(IParseTree tree)
        {
            return _summaryCoverageFactory.Create(Tokens.Variant);
        }

        public ISummaryCoverage VisitChildren(IRuleNode node)
        {
            var summary = _summaryCoverageFactory.Create(_typeName);
            foreach (var rangeClause in _caseClauseCtxt.rangeClause())
            {
                var inspRange = new UnreachableCaseInspectionRange(rangeClause);
                summary = VisitRange(inspRange, summary);
            }
            return summary;
        }

        private ISummaryCoverage VisitRange(UnreachableCaseInspectionRange inspRange, ISummaryCoverage summary)
        {
            try
            {
                if (inspRange.IsValueRange)
                {
                    var startCtxt = inspRange.GetChild<VBAParser.SelectStartValueContext>();
                    var start = startCtxt.Accept(_visitor);

                    var endCtxt = inspRange.GetChild<VBAParser.SelectEndValueContext>();
                    var end = endCtxt.Accept(_visitor);

                    summary.AddValueRange(start, end);
                    return summary;
                }

                var value = inspRange.Accept(_visitor);
                if (inspRange.IsLTorGT)
                {
                    var compOpCtxt = inspRange.GetChild<VBAParser.ComparisonOperatorContext>();
                    summary.AddIsClause(value, compOpCtxt.GetText());
                }
                else if (inspRange.IsRelationalOp)
                {
                    var relOp = inspRange.GetChild<VBAParser.RelationalOpContext>();
                    IUnreachableCaseInspectionValue LHS = null;
                    IUnreachableCaseInspectionValue RHS = null;
                    string opSymbol = string.Empty;
                    foreach (var ctxt in relOp.children.Where(ch => !(ch is VBAParser.WhiteSpaceContext)))
                    {
                        if((LHS is null) && ctxt is ParserRuleContext)
                        {
                            LHS = ctxt.Accept(_visitor);
                        }
                        else if ((RHS is null) && ctxt is ParserRuleContext)
                        {
                            RHS = ctxt.Accept(_visitor);
                        }
                        else
                        {
                            opSymbol = ctxt.GetText();
                        }
                    }
                    if(LHS.IsConstantValue && RHS.IsConstantValue && UnreachableCaseInspectionValueVisitor.BinaryOps.ContainsKey(opSymbol))
                    {
                        var op = UnreachableCaseInspectionValueVisitor.BinaryOps[opSymbol];
                        var result = op.Evaluate(LHS, RHS, summary.TypeName);
                        summary.AddSingleValue(result);
                    }
                    else
                    {
                        summary.AddRelationalOp(value);
                    }
                }
                else if (inspRange.IsSingleValue)
                {
                    summary.AddSingleValue(value);
                }
            }
            catch (ArgumentException)
            {
                IncompatibleRangeClauseDetectedArgs args = new IncompatibleRangeClauseDetectedArgs()
                {
                    RangeClause = (VBAParser.RangeClauseContext)inspRange.Context
                };
                OnIncompatibleRangeClauseDetected(args);
            }
            return summary;
        }

        public ISummaryCoverage VisitErrorNode(IErrorNode node)
        {
            return _summaryCoverageFactory.Create(Tokens.Variant);
        }

        public ISummaryCoverage VisitTerminal(ITerminalNode node)
        {
            return _summaryCoverageFactory.Create(Tokens.Variant);
        }

        protected virtual void OnIncompatibleRangeClauseDetected(IncompatibleRangeClauseDetectedArgs e)
        {
            IncompatibleRangeClauseDetected?.Invoke(this, e);
        }

        public event EventHandler<IncompatibleRangeClauseDetectedArgs> IncompatibleRangeClauseDetected;
    }

    public class IncompatibleRangeClauseDetectedArgs : EventArgs
    {
        public VBAParser.RangeClauseContext RangeClause { get; set; }
    }
}
