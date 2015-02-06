using System.Linq;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.Nodes
{
    public class OptionNode : Node
    {
        public enum VBOption
        {
            Base,
            Compare,
            Explicit,
            PrivateModule
        }

        public OptionNode(VisualBasic6Parser.OptionBaseStmtContext context, string scope)
            : base(context, scope)
        {
            _option = VBOption.Explicit;
            _value = null;
        }

        public OptionNode(VisualBasic6Parser.OptionCompareStmtContext context, string scope)
            : base(context, scope)
        {
            _option = VBOption.PrivateModule;
            _value = context.children.Last().GetText(); // note: context.children is a public field
        }

        public OptionNode(VisualBasic6Parser.OptionExplicitStmtContext context, string scope)
            : base(context, scope)
        {
            _option = VBOption.Explicit;
            _value = null;
        }

        public OptionNode(VisualBasic6Parser.OptionPrivateModuleStmtContext context, string scope)
            : base(context, scope)
        {
            _option = VBOption.PrivateModule;
            _value = null;
        }

        private readonly VBOption _option;
        public VBOption Option { get {  return _option; } }

        private readonly string _value;
        public string Value { get { return _value; } }
    }
}
