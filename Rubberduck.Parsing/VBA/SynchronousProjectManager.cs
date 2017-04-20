using System.Collections.Generic;
using System.Linq;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.VBA
{
    public class SynchronousProjectManager : ProjectManagerBase
    {
        public SynchronousProjectManager(
            RubberduckParserState state,
            IVBE vbe)
        :base(state, 
            vbe)
        { }


        public override IReadOnlyCollection<QualifiedModuleName> AllModules()
        {
            return Projects.SelectMany(project => project.VBComponents)
                            .Select(component => new QualifiedModuleName(component))
                            .ToHashSet()
                            .AsReadOnly(); ;
        }
    }
}
