using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.VBA
{
    public class SynchronousStateProjectManager : StateProjectManagerBase
    {
        public SynchronousStateProjectManager(
            RubberduckParserState state,
            IVBE vbe)
        :base(state, 
            vbe)
        { }


        public override IReadOnlyCollection<QualifiedModuleName> AllModules()
        {
            var modules = new HashSet<QualifiedModuleName>();
            foreach(var project in Projects.Select(tpl => tpl.Project))
            {
                using (var components = project.VBComponents)
                {
                    foreach (var component in components)
                    {
                        using (component)
                        {
                            modules.Add(component.QualifiedModuleName);
                        }
                    }
                }
            }
            return modules;
        }
    }
}
