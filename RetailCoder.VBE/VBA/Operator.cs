using System.Runtime.InteropServices;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA
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

    }
}
