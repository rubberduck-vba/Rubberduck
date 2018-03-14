using Rubberduck.Parsing.Grammar;
using Rubberduck.UI;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public class ExtractedParameter
    {
        public enum PassedBy
        {
            ByRef,
            ByVal
        }

        public static readonly string None = RubberduckUI.ExtractMethod_OutputNone;

        private readonly string _name;
        private readonly string _typeName;
        private readonly PassedBy _passedBy;

        public ExtractedParameter(string typeName, PassedBy passedBy, string name = null)
        {
            _name = name ?? None;
            _typeName = typeName;
            _passedBy = passedBy;
        }

        public string Name
        {
            get { return _name; }
        }

        public string TypeName
        {
            get { return _typeName; }
        }

        public PassedBy Passed
        {
            get { return _passedBy; }
        }

        public override string ToString()
        {
            return _passedBy.ToString() + ' ' + Name + ' ' + Tokens.As + ' ' + TypeName;
        }
    }
}
