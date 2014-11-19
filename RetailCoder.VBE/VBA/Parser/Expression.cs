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

        public override bool Equals(object obj)
        {
            return obj is Expression && ((Expression) obj).Value == _value;
        }

        public override int GetHashCode()
        {
            return _value.GetHashCode();
        }
    }
}
