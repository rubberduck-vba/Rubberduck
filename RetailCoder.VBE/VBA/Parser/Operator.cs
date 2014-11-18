using System.Runtime.InteropServices;
using Rubberduck.VBA.Parser.Grammar;

namespace Rubberduck.VBA.Parser
{
    [ComVisible(false)]
    public struct Operator
    {
        public Operator(string token, OperatorType type)
        {
            _token = token;
            _type = type;
        }

        private readonly string _token;
        public string Token { get { return _token; } }

        private readonly OperatorType _type;
        public OperatorType Type { get { return _type; } }

        public static Operator[] Operators()
        {
            return new[]
            {
                new Operator(@"^", OperatorType.MathOperator),
                new Operator(@"*", OperatorType.MathOperator),
                new Operator(@"/", OperatorType.MathOperator),
                new Operator(@"\", OperatorType.MathOperator),
                new Operator(@"Mod", OperatorType.MathOperator),
                new Operator(@"+", OperatorType.MathOperator),
                new Operator(@"-", OperatorType.MathOperator),
                new Operator(@"&", OperatorType.StringOperator),
                new Operator(@"=", OperatorType.ComparisonOperator), 
                new Operator(@"<>", OperatorType.ComparisonOperator),
                new Operator(@"<", OperatorType.ComparisonOperator),
                new Operator(@">", OperatorType.ComparisonOperator),
                new Operator(@"<=", OperatorType.ComparisonOperator), 
                new Operator(@">=", OperatorType.ComparisonOperator), 
                new Operator(@"Is", OperatorType.ComparisonOperator), 
                new Operator(ReservedKeywords.And, OperatorType.LogicalOperator),
                new Operator(ReservedKeywords.Or, OperatorType.LogicalOperator),
                new Operator(ReservedKeywords.XOr, OperatorType.LogicalOperator), 
                new Operator(ReservedKeywords.Not, OperatorType.LogicalOperator) 
            };
        }
    }
}
