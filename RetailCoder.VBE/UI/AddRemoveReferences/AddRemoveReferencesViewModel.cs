using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.AddRemoveReferences;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.UI.AddRemoveReferences
{
    public class AddRemoveReferencesViewModel : ViewModelBase
    {
        public AddRemoveReferencesViewModel(IEnumerable<ReferenceModel> model)
        {
            ComLibraries = model.Where(item => item.Type == ReferenceKind.TypeLibrary);
        }

        public IEnumerable<ReferenceModel> ComLibraries { get; }
    }
}
