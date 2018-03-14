using Rubberduck.Parsing.Grammar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Inspections.Concrete.UnreachableSelectCase
{
    public interface IUnreachableCaseInspectionBinaryOp
    {
        IUnreachableCaseInspectionValue Evaluate(IUnreachableCaseInspectionValue LHS, IUnreachableCaseInspectionValue RHS, string resultTypeName);
    }

    internal class UnreachableCaseInspectionBinaryOp : IUnreachableCaseInspectionBinaryOp
    {
        private string OpSymbol { set; get; }
        private Func<double, double, double> TheMathOp { set; get; }
        private Func<double, double, bool> TheLogicOp { set; get; }

        public UnreachableCaseInspectionBinaryOp(string opSymbol)
        {
            OpSymbol = opSymbol;
            if (MathOps.ContainsKey(opSymbol))
            {
                TheMathOp = MathOps[opSymbol];
            }
            else if (LogicOps.ContainsKey(opSymbol))
            {
                TheLogicOp = LogicOps[opSymbol];
            }
        }
        public virtual IUnreachableCaseInspectionValue Evaluate(IUnreachableCaseInspectionValue LHS, IUnreachableCaseInspectionValue RHS, string resultTypeName)
        {
            var cLHS = new UnreachableCaseInspectionValueConformed(LHS, LHS.TypeName);
            var cRHS = new UnreachableCaseInspectionValueConformed(RHS, RHS.TypeName);
            var mathResultType = resultTypeName ?? string.Empty;
            if (mathResultType.Equals(string.Empty))
            {
                mathResultType = DetermineMathResultType(cLHS, cRHS);
            }

            return CalculateResult(cLHS, cRHS, mathResultType);
        }

        protected static Dictionary<string, Func<double, double, double>> MathOps = new Dictionary<string, Func<double, double, double>>()
        {
            [UnreachableCaseInspectionValueVisitor.MathTokens.MULT] = delegate (double LHS, double RHS) { return LHS * RHS; },
            [UnreachableCaseInspectionValueVisitor.MathTokens.DIV] = delegate (double LHS, double RHS) { return LHS / RHS; },
            [UnreachableCaseInspectionValueVisitor.MathTokens.ADD] = delegate (double LHS, double RHS) { return LHS + RHS; },
            [UnreachableCaseInspectionValueVisitor.MathTokens.SUBTRACT] = delegate (double LHS, double RHS) { return LHS - RHS; },
            [UnreachableCaseInspectionValueVisitor.MathTokens.POW] = Math.Pow,
            [UnreachableCaseInspectionValueVisitor.MathTokens.MOD] = delegate (double LHS, double RHS) { return LHS % RHS; },
        };

        protected static Dictionary<string, Func<double, double, bool>> LogicOps = new Dictionary<string, Func<double, double, bool>>()
        {
            [CompareTokens.EQ] = delegate (double LHS, double RHS) { return LHS == RHS; },
            [CompareTokens.NEQ] = delegate (double LHS, double RHS) { return LHS != RHS; },
            [CompareTokens.LT] = delegate (double LHS, double RHS) { return LHS < RHS; },
            [CompareTokens.LTE] = delegate (double LHS, double RHS) { return LHS <= RHS; },
            [CompareTokens.GT] = delegate (double LHS, double RHS) { return LHS > RHS; },
            [CompareTokens.GTE] = delegate (double LHS, double RHS) { return LHS >= RHS; },
            [Tokens.And] = delegate (double LHS, double RHS) { return Convert.ToBoolean(LHS) && Convert.ToBoolean(RHS); },
            [Tokens.Or] = delegate (double LHS, double RHS) { return Convert.ToBoolean(LHS) || Convert.ToBoolean(RHS); },
            [Tokens.XOr] = delegate (double LHS, double RHS) { return Convert.ToBoolean(LHS) ^ Convert.ToBoolean(RHS); },
        };

        private IUnreachableCaseInspectionValue CalculateResult(IUnreachableCaseInspectionValue LHS, IUnreachableCaseInspectionValue RHS, string mathResultType)
        {
            var operands = GetOperandsAsDoubles(LHS, RHS);
            if (operands.Item1.HasValue && operands.Item2.HasValue)
            {
                if (!(TheMathOp is null))
                {
                    return UCIValueConverter.ConvertToType(TheMathOp(operands.Item1.Value, operands.Item2.Value), mathResultType);
                }
                else //logic operation
                {
                    return UCIValueConverter.ConvertToType(TheLogicOp(operands.Item1.Value, operands.Item2.Value), mathResultType);
                }
            }
            else 
            {
                LHS.IsConstantValue = operands.Item1.HasValue;
                RHS.IsConstantValue = operands.Item2.HasValue;
                var resultType = TheMathOp is null ? Tokens.Boolean : mathResultType;

                var result = new UnreachableCaseInspectionValue($"{LHS.ValueText} {OpSymbol} {RHS.ValueText}", resultType)
                {
                    IsConstantValue = false
                };
                return result;
            }
        }

        private Tuple<double?, double?> GetOperandsAsDoubles(IUnreachableCaseInspectionValue LHS, IUnreachableCaseInspectionValue RHS)
        {
            Tuple<double?, double?> result = null;
            try
            {
                double? nLHS = UCIValueConverter.ConvertDouble(LHS);
                double? nRHS = UCIValueConverter.ConvertDouble(RHS);
                result = new Tuple<double?, double?>(nLHS, nRHS);
            }
            catch (ArgumentException)
            {
                result = new Tuple<double?, double?>(null, null);
            }
            return result;
        }

        private string DetermineMathResultType(IUnreachableCaseInspectionValue LHS, IUnreachableCaseInspectionValue RHS)
        {
            var targetType = string.Empty;
            if (LHS.TypeName.Equals(RHS.TypeName))
            {
                targetType = LHS.TypeName;
            }
            else if (LHS.TypeName.Equals(Tokens.Currency) || RHS.TypeName.Equals(Tokens.Currency))
            {
                targetType = Tokens.Currency;
            }
            else if (LHS.TypeName.Equals(Tokens.Double) || RHS.TypeName.Equals(Tokens.Double))
            {
                targetType = Tokens.Double;
            }
            else
            {
                targetType = LHS.TypeName;
            }
            return targetType;
        }
    }
}
