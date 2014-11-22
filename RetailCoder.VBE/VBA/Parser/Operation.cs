using System.Runtime.InteropServices;

namespace Rubberduck.VBA.Parser
{
    [ComVisible(false)]
    public struct Operation
    {
        public Operation(Expression left, string @operator, Expression right, OperatorType type)
        {
            _left = left;
            _operator = @operator;
            _right = right;
            _type = type;
        }

        private readonly Expression _left;
        /// <summary>
        /// Gets the left operand of the operation.
        /// </summary>
        public Expression LeftOperand { get { return _left; } }

        private readonly string _operator;
        /// <summary>
        /// Gets 
        /// </summary>
        public string Operator { get { return _operator; } }

        private readonly Expression _right;
        /// <summary>
        /// Gets the left operand of the operation.
        /// </summary>
        public Expression RightOperand { get { return _right; } }

        private readonly OperatorType _type;
        /// <summary>
        /// Gets the type of operator used in the operation.
        /// </summary>
        public OperatorType OperatorType { get { return _type; } }
    }
}