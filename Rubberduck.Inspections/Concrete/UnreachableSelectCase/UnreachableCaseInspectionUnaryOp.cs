using Rubberduck.Parsing.Grammar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Inspections.Concrete.UnreachableSelectCase
{
    public interface IUnreachableCaseInspectionUnaryOp
    {
        IUnreachableCaseInspectionValue Evaluate(IUnreachableCaseInspectionValue value, string resultTypeName);
        string OpSymbol { set; get; }
        Func<double, double> TheMathOp { set; get; }
        Func<double, bool> TheLogicOp { set; get; }
    }

    internal class UnreachableCaseInspectionUnaryOp : IUnreachableCaseInspectionUnaryOp
    {
        public UnreachableCaseInspectionUnaryOp(string opSymbol)
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
        public string OpSymbol { set; get; }
        public Func<double, double> TheMathOp { set; get; }
        public Func<double, bool> TheLogicOp { set; get; }

        public virtual IUnreachableCaseInspectionValue Evaluate(IUnreachableCaseInspectionValue value, string resultTypeName)
        {
            var targetType = resultTypeName ?? string.Empty;
            if (targetType.Equals(string.Empty))
            {
                targetType = value.TypeName;
            }

            return CalculateResult(value, targetType);
        }

        private IUnreachableCaseInspectionValue CalculateResult(IUnreachableCaseInspectionValue value, string targetType)
        {
            SafeConvertToNullable<double>(value, out double? operand);
            if (operand.HasValue)
            {
                if (!(TheMathOp is null))
                {
                    return UCIValueConverter.ConvertToType(TheMathOp(operand.Value), targetType);
                }
                else
                {
                    return UCIValueConverter.ConvertToType(TheLogicOp(operand.Value), targetType);
                }
            }
            else
            {
                value.IsConstantValue = operand.HasValue;
                var result = new UnreachableCaseInspectionValue($"{OpSymbol} {value.ValueText}", targetType)
                {
                    IsConstantValue = false
                };
                return result;
            }
        }

        private void SafeConvertToNullable<T>(IUnreachableCaseInspectionValue value, out double? convertedValue)
        {
            try
            {
                convertedValue = UCIValueConverter.ConvertDouble(value);
            }
            catch (ArgumentException)
            {
                convertedValue = null;
            }
        }

        protected Dictionary<string, Func<double, double>> MathOps = new Dictionary<string, Func<double, double>>()
        {
            [UnreachableCaseInspectionValueVisitor.MathTokens.ADDITIVE_INVERSE] = delegate (double value) { return value * -1.0; }
        };

        protected Dictionary<string, Func<double, bool>> LogicOps = new Dictionary<string, Func<double, bool>>()
        {
            [Tokens.Not] = delegate (double value) { return !(Convert.ToBoolean(value)); }
        };
    }

    //internal class UnreachableCaseInspectionNot : UnreachableCaseInspectionUnaryOp
    //{
    //    public UnreachableCaseInspectionNot()
    //    {
    //        OpSymbol = "Not";
    //        TheLogicOp = delegate (double value) { return !(Convert.ToBoolean(value)); };
    //    }
    //}

    //internal class UnreachableCaseInspectionMinusOp : UnreachableCaseInspectionUnaryOp
    //{
    //    public UnreachableCaseInspectionMinusOp()
    //    {
    //        OpSymbol = "-";
    //        TheMathOp = delegate (double value) { return value * -1.0; };
    //    }
    //}

}
