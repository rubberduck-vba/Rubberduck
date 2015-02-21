using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Inspections;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.ParseTreeListeners
{
    /// <summary>
    /// This class is not used, because the grammar (/generated parser)
    /// requires options to be specified first, or Module options end up in an error node.
    /// </summary>
    public class ModuleOptionsListener : VBListenerBase, IExtensionListener<VBParser.ModuleOptionContext>
    {
        private readonly QualifiedModuleName _qualifiedName;

        private readonly IList<QualifiedContext<VBParser.ModuleOptionContext>> _members =
            new List<QualifiedContext<VBParser.ModuleOptionContext>>();

        public ModuleOptionsListener(QualifiedModuleName qualifiedName)
        {
            _qualifiedName = qualifiedName;
        }

        public IEnumerable<QualifiedContext<VBParser.ModuleOptionContext>> Members { get { return _members; } }


        public override void EnterModuleOptions(VBParser.ModuleOptionsContext context)
        {
            foreach (var option in context.ModuleOption())
            {
                _members.Add(new QualifiedContext<VBParser.ModuleOptionContext>(_qualifiedName, option));
            }
        }
    }
}
