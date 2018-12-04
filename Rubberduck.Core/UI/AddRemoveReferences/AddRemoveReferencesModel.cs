using System.Collections.Generic;
using System.Linq;
using Rubberduck.AddRemoveReferences;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Settings;

namespace Rubberduck.UI.AddRemoveReferences
{
    public interface IAddRemoveReferencesModel
    {
        IGeneralSettings Settings { get; set; }
        ProjectDeclaration Project { get; set; }
        List<ReferenceModel> References { get; set; }
        IReadOnlyList<ReferenceModel> NewReferences { get; set; }
    }

    public class AddRemoveReferencesModel : IAddRemoveReferencesModel
    {
        public AddRemoveReferencesModel(ProjectDeclaration project, IEnumerable<ReferenceModel> references, IGeneralSettings settings)
        {
            Settings = settings;
            Project = project;
            References = references.ToList();
        }

        public IGeneralSettings Settings { get; set; }

        public ProjectDeclaration Project { get; set; }

        public List<ReferenceModel> References { get; set; }

        public IReadOnlyList<ReferenceModel> NewReferences { get; set; }
    }
}
