using Rubberduck.Extensions;

namespace Rubberduck.VBA.Nodes
{
    public class ProcedureNode : Node
    {
        public enum VBProcedureKind
        {
            Sub,
            Function,
            PropertyGet,
            PropertyLet,
            PropertySet
        }

        public ProcedureNode(Selection selection, string project, string module, VBProcedureKind kind, string name, string returnType = null, VBAccessibility accessibility = VBAccessibility.Public)
            : base(selection, project, module)
        {
            _name = name;
            _returnType = returnType;
            _kind = kind;
            _accessibility = accessibility;
        }

        private readonly string _name;
        public string Name { get { return _name; } }

        private readonly string _returnType;
        public string ReturnType { get { return _returnType; } }

        private readonly VBProcedureKind _kind;
        public VBProcedureKind Kind { get { return _kind; } }

        private readonly VBAccessibility _accessibility;
        public VBAccessibility Accessibility { get { return _accessibility; } }
    }
}