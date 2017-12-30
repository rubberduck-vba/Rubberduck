using System.Collections.Generic;
using System.Linq;
using Rubberduck.UI;

namespace Rubberduck.AddRemoveReferences
{
    public class AddRemoveReferencesViewModel : ViewModelBase
    {
        public AddRemoveReferencesViewModel(IReadOnlyList<ReferenceModel> model)
        {
            AvailableReferences = model;
        }

        public IEnumerable<ReferenceModel> AvailableReferences { get; }
    }
}
