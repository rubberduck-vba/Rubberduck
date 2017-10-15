using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using System.Threading.Tasks;
using System.Collections.Concurrent;

namespace Rubberduck.Parsing.VBA
{
    public class ProjectManager : ProjectManagerBase
    {
        private const int _maxDegreeOfQMNCreationParallelism = -1;

        public ProjectManager(
            RubberduckParserState state,
            IVBE vbe)
        :base(state, 
            vbe)
        { }


        public override IReadOnlyCollection<QualifiedModuleName> AllModules()
        {
            var components = Projects.SelectMany(project => project.VBComponents);

            var options = new ParallelOptions();
            options.MaxDegreeOfParallelism = _maxDegreeOfQMNCreationParallelism;

            var modules = new ConcurrentBag<QualifiedModuleName>();
            Parallel.ForEach(components,
                options,
                component =>
                {
                    modules.Add(new QualifiedModuleName(component));
                }
            );

            return modules.ToHashSet().AsReadOnly();
        }
    }
}
