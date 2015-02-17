using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.Nodes
{
    public class ParameterNode : Node
    {
        public enum VBParameterType
        {
            ImplicitByRef,
            ByRef,
            ByVal
        }

        public ParameterNode(VBParser.ArgContext context, string scope)
            : base(context, scope)
        {
        }

        private readonly VBParameterType _passedBy;
        public VBParameterType PassedBy { get { return _passedBy; } }

        private readonly string _name;
        public string Name { get { return _name; } }

        private readonly string _type;
        public string TypeName { get { return _type; } }

        private readonly bool _isOptional;
        public bool IsOptional { get { return _isOptional; } }
    }
}