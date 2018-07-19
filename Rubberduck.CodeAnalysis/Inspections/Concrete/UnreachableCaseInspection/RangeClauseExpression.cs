using Rubberduck.Parsing.Grammar;
using System.Collections.Generic;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IRangeClauseExpression
    {
        IParseTreeValue RHSValue { get; }
        IParseTreeValue LHSValue { get; }
        string LHS { get; }
        string RHS { get; }
        string OpSymbol { get; }
        bool IsMismatch { set; get; }
        bool IsUnreachable { set; get; }
        bool IsInherentlyUnreachable { set; get; }
    }

    public class RangeOfValuesExpression : RangeClauseExpression
    {
        public RangeOfValuesExpression((IParseTreeValue lhs, IParseTreeValue rhs) rangeOfValues)
            : base(rangeOfValues.lhs, rangeOfValues.rhs, Tokens.To) { }
    }

    public class BinaryExpression : RangeClauseExpression
    {
        public BinaryExpression(IParseTreeValue lhs, IParseTreeValue rhs, string opSymbol)
            : base(lhs, rhs, opSymbol, true)
        {
            _hashCode = OpSymbol.GetHashCode();
        }
    }

    public class LikeExpression : RangeClauseExpression
    {
        public LikeExpression(IParseTreeValue lhs, IParseTreeValue rhs)
            : base(lhs, rhs, Tokens.Like, false)
        {
            _hashCode = OpSymbol.GetHashCode();
        }
        public string Operand => LHS;
        public string Pattern => RHS;
        public bool Filters(LikeExpression like)
        {
            //TODO: Enhancement - evaluate Like Pattern for superset/subset conditions.
            //e.g., "*" would filter "?*", or "?*" would filter "a*" 
            //They go here...
            if ( like.Operand.Equals(Operand) && Pattern.Equals("*"))//The easiest one
            {
                return true;
            }
            return false;
        }
    }

    public class IsClauseExpression : RangeClauseExpression
    {
        public IsClauseExpression(IParseTreeValue value, string opSymbol)
            : base(value, null, opSymbol)
        {
            _hashCode = OpSymbol.GetHashCode();
        }

        public override string ToString()
        {
            return $"Is {OpSymbol} {LHSValue}";
        }
    }

    public class UnaryExpression : RangeClauseExpression
    {
        public UnaryExpression(IParseTreeValue value, string opSymbol)
            : base(value, null, opSymbol)
        {
            _hashCode = ToString().GetHashCode();
        }

        public override string ToString()
        {
            return $"{OpSymbol} {LHSValue}";
        }
    }

    public class ValueExpression : RangeClauseExpression
    {
        public ValueExpression(IParseTreeValue value)
            : base(value, null, string.Empty)
        {
            _hashCode = LHS.GetHashCode();
        }

        public override string ToString() => LHS;
    }

    public abstract class RangeClauseExpression : IRangeClauseExpression
    {
        private ClauseExpressionData _data;
        protected int _hashCode;

        public IParseTreeValue LHSValue => _data.LHSValue;
        public IParseTreeValue RHSValue => _data.RHSValue;
        public string LHS => _data.LHS;
        public string RHS => _data.RHS;
        public string OpSymbol => _data.OpSymbol;
        public bool IsMismatch { set => _data.IsMismatch = value; get => _data.IsMismatch; }
        public bool IsUnreachable { set => _data.IsUnreachable = value; get => _data.IsUnreachable; }
        public bool IsInherentlyUnreachable { set => _data.IsInherentlyUnreachable = value; get => _data.IsInherentlyUnreachable; }

        public RangeClauseExpression(IParseTreeValue lhs, IParseTreeValue rhs, string opSymbol, bool sortOperands = false)
        {
            _data = new ClauseExpressionData(lhs, rhs, opSymbol);
            if (sortOperands)
            {
                SortExpressionOperands();
            }
            _hashCode = ToString().GetHashCode();
        }

        public override int GetHashCode()
        {
            return _hashCode;
        }

        public override bool Equals(object obj)
        {
            if (!(obj is RangeClauseExpression expression))
            {
                return false;
            }

            return ToString().Equals(expression.ToString());
        }

        public override string ToString()
        {
            return $"{LHS} {OpSymbol} {RHS}";
        }

        private void SortExpressionOperands()
        {
            if ((LHSValue.ParsesToConstantValue && !RHSValue.ParsesToConstantValue
                || !LHSValue.ParsesToConstantValue && !RHSValue.ParsesToConstantValue && LHS.CompareTo(RHS) > 0)
                && AlgebraicInverses.ContainsKey(OpSymbol))
            {
                var lhs = RHSValue;
                var rhs = LHSValue;
                _data = new ClauseExpressionData(lhs, rhs, AlgebraicInverses[OpSymbol]);
            }
        }

        private static Dictionary<string, string> AlgebraicInverses = new Dictionary<string, string>()
        {
            [RelationalOperators.LT] = RelationalOperators.GT,
            [RelationalOperators.NEQ] = RelationalOperators.NEQ,
            [LogicalOperators.AND] = LogicalOperators.AND,
            [LogicalOperators.OR] = LogicalOperators.OR,
            [LogicalOperators.XOR] = LogicalOperators.XOR,
            [RelationalOperators.LTE] = RelationalOperators.GTE,
            [RelationalOperators.LTE2] = RelationalOperators.GTE,
            [RelationalOperators.GT] = RelationalOperators.LT,
            [RelationalOperators.GTE] = RelationalOperators.LTE,
            [RelationalOperators.GTE2] = RelationalOperators.LTE,
            [RelationalOperators.EQ] = RelationalOperators.EQ,
        };

        private struct ClauseExpressionData : IRangeClauseExpression
        {
            public IParseTreeValue LHSValue { private set; get; }
            public IParseTreeValue RHSValue { private set; get; }
            public string LHS { private set; get; }
            public string RHS { private set; get; }
            public string OpSymbol { private set; get; }
            public bool IsMismatch { set; get; }
            public bool IsUnreachable { set; get; }
            public bool IsInherentlyUnreachable { set; get; }

            public ClauseExpressionData(IParseTreeValue lhs, IParseTreeValue rhs, string opSymbol)
            {
                RHSValue = rhs;
                RHS = rhs is null ? string.Empty : rhs.ValueText;
                LHSValue = lhs;
                LHS = lhs is null ? string.Empty : lhs.ValueText;
                OpSymbol = opSymbol;
                IsMismatch = false;
                IsUnreachable = false;
                IsInherentlyUnreachable = false;
            }
        }
    }
}
