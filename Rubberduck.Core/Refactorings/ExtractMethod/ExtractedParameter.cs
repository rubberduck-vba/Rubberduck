using Rubberduck.Parsing.Grammar;
using Rubberduck.Resources;
using Tokens = Rubberduck.Resources.Tokens;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public class ExtractedParameter
    {
        public enum PassedBy
        {
            ByRef,
            ByVal
        }

        public static readonly string None = RefactoringsUI.ExtractMethod_OutputNone;

        public ExtractedParameter(string typeName, PassedBy passedBy, string name = null)
        {
            Name = name ?? None;
            TypeName = typeName;
            Passed = passedBy;
        }

        public string Name { get; }

        public string TypeName { get; }

        public PassedBy Passed { get; }

        public override string ToString()
        {
            return Passed.ToString() + ' ' + Name + ' ' + Tokens.As + ' ' + TypeName;
        }
    }
}
