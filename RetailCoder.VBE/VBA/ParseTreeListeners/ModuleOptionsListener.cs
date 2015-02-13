using System.Collections.Generic;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.VBA.ParseTreeListeners
{
    /// <summary>
    /// This class is not used, because the grammar (/generated parser)
    /// requires options to be specified first, or module options end up in an error node.
    /// </summary>
    public class ModuleOptionsListener : VisualBasic6BaseListener, IExtensionListener<VisualBasic6Parser.ModuleOptionContext>
    {
        private readonly IList<VisualBasic6Parser.ModuleOptionContext> _members = new List<VisualBasic6Parser.ModuleOptionContext>();
        public IEnumerable<VisualBasic6Parser.ModuleOptionContext> Members { get { return _members; } }

        public override void EnterModuleOptions(VisualBasic6Parser.ModuleOptionsContext context)
        {
            foreach (var option in context.moduleOption())
            {
                _members.Add(option);
            }
        }
    }
}