using Rubberduck.VBA.Grammar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

        public OptionNode(VBParser.OptionBaseStmtContext context, string scope)
            : base(context, scope)
        {
            _option = VBOption.Explicit;
            _value = null;
        }

        public OptionNode(VBParser.OptionCompareStmtContext context, string scope)
            : base(context, scope)
        {
            _option = VBOption.PrivateModule;
            _value = context.children.Last().GetText(); // note: context.children is a public field
        }

        public OptionNode(VBParser.OptionExplicitStmtContext context, string scope)
            : base(context, scope)
        {
            _option = VBOption.Explicit;
            _value = null;
        }

        public OptionNode(VBParser.OptionPrivateModuleStmtContext context, string scope)
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
