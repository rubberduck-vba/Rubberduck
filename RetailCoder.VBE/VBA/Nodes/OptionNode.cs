using System.Linq;
using Rubberduck.Parsing.Grammar;

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

        public OptionNode(VBAParser.OptionBaseStmtContext context, string scope)
            : base(context, scope)
        {
            _option = VBOption.Explicit;
            _value = null;
        }

        public OptionNode(VBAParser.OptionCompareStmtContext context, string scope)
            : base(context, scope)
        {
            _option = VBOption.PrivateModule;
            _value = context.children.Last().GetText(); // note: context.children is a public field
        }

        public OptionNode(VBAParser.OptionExplicitStmtContext context, string scope)
            : base(context, scope)
        {
            _option = VBOption.Explicit;
            _value = null;
        }

        public OptionNode(VBAParser.OptionPrivateModuleStmtContext context, string scope)
            : base(context, scope)
        {
            _option = VBOption.PrivateModule;
            _value = null;
        }

        private readonly VBOption _option;
        public VBOption Option { get { return _option; } }

        private readonly string _value;
        public string Value { get { return _value; } }
    }
}
