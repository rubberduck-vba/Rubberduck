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

        public ProcedureNode(VisualBasic6Parser.SubStmtContext context, string scope, string localScope)
            : base(context, scope, localScope)
        {
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