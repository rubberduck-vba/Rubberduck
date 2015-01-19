using Rubberduck.Extensions;

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

        public ParameterNode(Selection selection, string project, string module, VBParameterType passedBy, string name, string type, bool isOptional = false)
            : base(selection, project, module)
        {
            _passedBy = passedBy;
            _name = name;
            _type = type;
            _isOptional = isOptional;
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