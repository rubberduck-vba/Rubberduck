using Rubberduck.VBA;

namespace Rubberduck.UI.Refactorings.ExtractMethod
{
    public class ExtractedParameter
    {
        public enum PassedBy
        {
            ByRef,
            ByVal
        }

        public ExtractedParameter(string name, string typeName, PassedBy passed)
        {
            Name = name;
            TypeName = typeName;
            Passed = passed;
        }

        public string Name { get; set; }
        public string TypeName { get; set; }
        public PassedBy Passed { get; set; }

        public override string ToString()
        {
            return Passed.ToString() + ' ' + Name + ' ' + Tokens.As + ' ' + TypeName;
        }
    }
}