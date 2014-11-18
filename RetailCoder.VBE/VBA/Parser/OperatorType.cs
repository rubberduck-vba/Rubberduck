using System.Runtime.InteropServices;

namespace Rubberduck.VBA.Parser
{
    [ComVisible(false)]
    public enum OperatorType
    {
        MathOperator,
        StringOperator,
        ComparisonOperator,
        LogicalOperator
    }
}