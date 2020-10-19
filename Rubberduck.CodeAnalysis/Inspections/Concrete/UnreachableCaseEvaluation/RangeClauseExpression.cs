using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Refactorings;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete.UnreachableCaseEvaluation
{
    internal interface IRangeClauseExpression
    {
        IParseTreeValue RHS { get; }
        IParseTreeValue LHS { get; }
        string OpSymbol { get; }
        bool IsMismatch { set; get; }
        bool IsUnreachable { set; get; }
        bool IsOverflow { set; get; }
        bool IsInherentlyUnreachable { set; get; }
    }

    internal class RangeOfValuesExpression : RangeClauseExpression
    {
        public RangeOfValuesExpression((IParseTreeValue lhs, IParseTreeValue rhs) rangeOfValues)
            : base(rangeOfValues.lhs, rangeOfValues.rhs, Tokens.To) { }
    }

    internal class BinaryExpression : RangeClauseExpression
    {
        public BinaryExpression(IParseTreeValue lhs, IParseTreeValue rhs, string opSymbol)
            : base(lhs, rhs, opSymbol, true)
        {
            _hashCode = OpSymbol.GetHashCode();
        }
    }

    internal class LikeExpression : RangeClauseExpression
    {
        public LikeExpression(IParseTreeValue lhs, IParseTreeValue rhs)
            : base(lhs, rhs, Tokens.Like, false)
        {
            _hashCode = OpSymbol.GetHashCode();
        }
        public string Operand => LHS.Token;
        public string Pattern => AnnotateAsStringConstant(RHS.Token);
        public bool Filters(LikeExpression like)
        {
            //TODO: Enhancement - evaluate Like Pattern for superset/subset conditions.
            //e.g., "*" would filter "?*", or "?*" would filter "a*" 
            //They go here...
            if ( like.Operand.Equals(Operand) && Pattern.Equals($"\"*\""))//The easiest one
            {
                return true;
            }
            return false;
        }

        private static string AnnotateAsStringConstant(string input)
        {
            var result = input;
            if (!input.StartsWith("\""))
            {
                result = $"\"{result}";
            }
            if (!input.EndsWith("\""))
            {
                result = $"{result}\"";
            }
            return result;
        }
    }

    internal class IsClauseExpression : RangeClauseExpression
    {
        public IsClauseExpression(IParseTreeValue value, string opSymbol)
            : base(value, null, opSymbol)
        {
            _hashCode = OpSymbol.GetHashCode();
        }

        public override string ToString()
        {
            return $"Is {OpSymbol} {LHS}";
        }
    }

    internal class UnaryExpression : RangeClauseExpression
    {
        public UnaryExpression(IParseTreeValue value, string opSymbol)
            : base(value, null, opSymbol)
        {
            _hashCode = ToString().GetHashCode();
        }

        public override string ToString()
        {
            return $"{OpSymbol} {LHS}";
        }
    }

    internal class ValueExpression : RangeClauseExpression
    {
        public ValueExpression(IParseTreeValue value)
            : base(value, null, string.Empty)
        {
            _hashCode = LHS.Token.GetHashCode();
        }

        public override string ToString() => LHS.Token;
    }

    internal abstract class RangeClauseExpression : IRangeClauseExpression
    {
        private ClauseExpressionData _data;
        protected int _hashCode;

        public IParseTreeValue LHS => _data.LHS;
        public IParseTreeValue RHS => _data.RHS;
        public string OpSymbol => _data.OpSymbol;
        public bool IsMismatch { set => _data.IsMismatch = value; get => _data.IsMismatch; }
        public bool IsOverflow { set => _data.IsOverflow = value; get => _data.IsOverflow; }
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
            return $"{LHS.Token} {OpSymbol} {RHS.Token}";
        }

        private void SortExpressionOperands()
        {
            if ((LHS.ParsesToConstantValue && !RHS.ParsesToConstantValue
                || !LHS.ParsesToConstantValue && !RHS.ParsesToConstantValue && LHS.Token.CompareTo(RHS.Token) > 0)
                && AlgebraicInverses.ContainsKey(OpSymbol))
            {
                var lhs = RHS;
                var rhs = LHS;
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
            public IParseTreeValue LHS { private set; get; }
            public IParseTreeValue RHS { private set; get; }
            public string OpSymbol { private set; get; }
            public bool IsMismatch { set; get; }
            public bool IsOverflow { set; get; }
            public bool IsUnreachable { set; get; }
            public bool IsInherentlyUnreachable { set; get; }

            public ClauseExpressionData(IParseTreeValue lhs, IParseTreeValue rhs, string opSymbol)
            {
                RHS = rhs;
                LHS = lhs;
                OpSymbol = opSymbol;
                IsMismatch = false;
                IsOverflow = (LHS != null && LHS.IsOverflowExpression) || (RHS != null && RHS.IsOverflowExpression);
                IsUnreachable = false;
                IsInherentlyUnreachable = false;
            }
        }
    }
}
