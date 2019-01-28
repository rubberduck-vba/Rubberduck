using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.AddRemoveReferences;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;

namespace Rubberduck.UI.AddRemoveReferences
{
    public interface IAddRemoveReferencesModel
    {
        string HostApplication { get; }
        RubberduckParserState State { get; }
        ReferenceSettings Settings { get; set; }
        ProjectDeclaration Project { get; set; }
        List<ReferenceModel> References { get; set; }
        IReadOnlyList<ReferenceModel> NewReferences { get; set; }
    }

    public class AddRemoveReferencesModel : IAddRemoveReferencesModel
    {
        private static readonly string Host = Path.GetFileName(Application.ExecutablePath).ToUpperInvariant();

        public AddRemoveReferencesModel(RubberduckParserState state, ProjectDeclaration project, IEnumerable<ReferenceModel> references, ReferenceSettings settings)
        {
            State = state;
            Settings = settings;
            Project = project;
            References = references.ToList();
            try
            {
                ApplySettingsToModel();
            }
            catch
            {
                //Meh.
            }
        }

        public string HostApplication => Host;

        public RubberduckParserState State { get; }

        public ReferenceSettings Settings { get; set; }

        public ProjectDeclaration Project { get; set; }

        public List<ReferenceModel> References { get; set; }

        public IReadOnlyList<ReferenceModel> NewReferences { get; set; }

        private void ApplySettingsToModel()
        {
            var recent = Settings.GetRecentReferencesForHost(HostApplication);
            foreach (var item in recent)
            {
                var match = References.FirstOrDefault(reference => reference.Matches(item));
                if (match != null)
                {
                    match.IsRecent = true;
                }
            }

            var pinned = Settings.GetPinnedReferencesForHost(HostApplication);
            foreach (var item in pinned)
            {
                var match = References.FirstOrDefault(reference => reference.Matches(item));
                if (match != null)
                {
                    match.IsPinned = true;
                }
            }
        }
    }
}
