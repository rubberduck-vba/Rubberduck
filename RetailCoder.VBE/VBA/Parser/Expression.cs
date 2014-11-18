using System.Runtime.InteropServices;

namespace Rubberduck.VBA.Parser
{
    [ComVisible(false)]
    public struct Expression
    {
        public Expression(string value)
        {
            _value = value;
        }

        private readonly string _value;
        public string Value { get { return _value; } }
    }
}
