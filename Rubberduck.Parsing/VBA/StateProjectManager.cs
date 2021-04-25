using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System.Threading.Tasks;
using System.Collections.Concurrent;
using Rubberduck.InternalApi.Extensions;
using Rubberduck.Parsing.VBA.Extensions;

namespace Rubberduck.Parsing.VBA
{
    public class StateProjectManager : StateProjectManagerBase
    {
        private const int _maxDegreeOfQMNCreationParallelism = -1;

        public StateProjectManager(
            RubberduckParserState state,
            IVBE vbe)
        :base(state, 
            vbe)
        { }


        public override IReadOnlyCollection<QualifiedModuleName> AllModules()
        {
            var options = new ParallelOptions();
            options.MaxDegreeOfParallelism = _maxDegreeOfQMNCreationParallelism;

            var modules = new ConcurrentBag<QualifiedModuleName>();
            foreach (var project in Projects.Select(tpl => tpl.Project))
            {
                using (var components = project.VBComponents)
                {
                    Parallel.ForEach(components,
                        options,
                        component =>
                        {
                            using (component)
                            {
                                modules.Add(component.QualifiedModuleName);
                            }
                        }
                    );
                }
            }

            return modules.ToHashSet().AsReadOnly();
        }
    }
}
