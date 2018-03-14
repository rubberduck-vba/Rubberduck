using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Antlr4.Runtime;
using Antlr4.Runtime.Tree;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Inspections.Concrete.UnreachableSelectCase
{
    public interface IUnreachableCaseInspectionRange : IUnreachableCaseInspectionContext
    {
        bool HasCoverage { get; }
        bool HasIncompatibleType { set; get; }
        bool IsValueRange { get; }
        bool IsLTorGT { get; }
        bool IsSingleValue { get; }
        bool IsRelationalOp { get; }
    }

    public class UnreachableCaseInspectionRange : UnreachableCaseInspectionContext, IUnreachableCaseInspectionRange
    {
        public UnreachableCaseInspectionRange(VBAParser.RangeClauseContext context) : base(context)
        {
        }

        public bool HasCoverage => IsValueRange || IsLTorGT || IsSingleValue || IsRelationalOp;
        public bool HasIncompatibleType { get; set; }
        public bool IsValueRange => Context.HasChildToken(Tokens.To);
        public bool IsLTorGT => Context.HasChildToken(Tokens.Is);
        public bool IsSingleValue => !(IsValueRange && IsLTorGT && IsRelationalOp);
        public bool IsRelationalOp => Context.TryGetChildContext<VBAParser.RelationalOpContext>(out _);
    }

    public interface IUnreachableCaseInspectionCaseClause : IUnreachableCaseInspectionContext
    {
        //List<IUnreachableCaseInspectionRange> Ranges { set; get; }
        //bool CanBeInspected { get; }
        bool IsIncompatibleType { get; }
    }

    public class UnreachableCaseInspectionCaseClause : UnreachableCaseInspectionContext, IUnreachableCaseInspectionCaseClause
    {
        private List<IUnreachableCaseInspectionRange> _ranges;
        private IParseTree _caseClause;
        public UnreachableCaseInspectionCaseClause(VBAParser.CaseClauseContext caseClause) : base(caseClause)
        {
            _caseClause = caseClause;
            _ranges = new List<IUnreachableCaseInspectionRange>();
            foreach (var range in caseClause.rangeClause())
            {
                _ranges.Add(new UnreachableCaseInspectionRange(range));
            }
        }

        //public void OnIncompatibleRangeClauseDetected(IncompatibleRangeClauseDetectedArgs e)
        //{
        //    IncompatibleRangeClauseDetected?.Invoke(this, e);
        //}

        //public List<IUnreachableCaseInspectionRange> Ranges { get; }
        //public bool CanBeInspected => _ranges.Any(rg => rg.HasCoverage);
        public bool IsIncompatibleType => _ranges.Any(rg => rg.HasIncompatibleType);
    }
}
