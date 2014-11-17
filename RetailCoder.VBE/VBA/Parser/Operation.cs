using System.Runtime.InteropServices;

namespace Rubberduck.VBA.Parser
{
    [ComVisible(false)]
    public struct Operation
    {
        public Operation(Expression left, string @operator, Expression right)
        {
            _left = left;
            _operator = @operator;
            _right = right;
        }

        private readonly Expression _left;
        public Expression LeftOperand { get { return _left; } }

        private readonly string _operator;
        public string Operator { get { return _operator; } }

        private readonly Expression _right;
        public Expression RightOperand { get { return _right; } }
    }
}