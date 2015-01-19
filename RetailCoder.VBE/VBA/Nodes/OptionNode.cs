using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Extensions;

namespace Rubberduck.VBA.Nodes
{
    public class OptionNode : Node
    {
        public enum VBOption
        {
            Base,
            Compare,
            Explicit
        }

        public OptionNode(Selection selection, string project, string module, VBOption option, string value = null)
            : base(selection, project, module)
        {
            _option = option;
            _value = value;
        }

        private readonly VBOption _option;
        public VBOption Option { get {  return _option; } }

        private readonly string _value;
        public string Value { get { return _value; } }
    }
}
