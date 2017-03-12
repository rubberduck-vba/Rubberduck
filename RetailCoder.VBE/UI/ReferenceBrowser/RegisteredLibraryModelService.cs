using System.Collections.Generic;
using System.Linq;
using Kavod.ComReflection;

namespace Rubberduck.UI.ReferenceBrowser
{
    public class RegisteredLibraryModelService
    {
        public List<RegisteredLibraryModel> GetAllRegisteredLibraries()
        {
            return (from r in LibraryRegistration.GetRegisteredTypeLibraryEntries()
                    select new RegisteredLibraryModel(r)
                    ).ToList();
        }
    }
}