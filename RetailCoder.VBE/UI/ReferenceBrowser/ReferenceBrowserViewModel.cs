using System;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Vbe.Interop;

namespace Rubberduck.UI.ReferenceBrowser
{
    public class ReferenceBrowserViewModel : ViewModelBase
    {
        private readonly VBE _vbe;

        public ReferenceBrowserViewModel(VBE vbe, RegisteredLibraryModelService service)
        {
            _vbe = vbe;
            
            ComReferences = new RegisteredComLibrariesViewModel(service, vbe);
            VbaProjectReferences = new VbaProjectLibrariesViewModel(vbe);
        }

        public LibrariesViewModel ComReferences { get; }

        public LibrariesViewModel VbaProjectReferences { get; }

        private class RegisteredComLibrariesViewModel : LibrariesViewModel
        {
            public RegisteredComLibrariesViewModel(RegisteredLibraryModelService service, VBE vbe)
            {
                Build(service, vbe);
            }

            private async void Build(RegisteredLibraryModelService service, VBE vbe)
            {
                var models = service.GetAllRegisteredLibraries();

                var list = models
                    .Select(l => new RegisteredLibraryViewModel(l, vbe.ActiveVBProject))
                    .ToList();

                await Task.Run(() => list.Sort());

                foreach (var vm in list)
                {
                    _registeredLibraries.Add(vm);
                }
            }
        }

        private class VbaProjectLibrariesViewModel : LibrariesViewModel
        {
            public VbaProjectLibrariesViewModel(VBE vbe)
            {
                foreach (var reference in vbe.ActiveVBProject.References.OfType<Reference>())
                {
                    if (reference.Type == vbext_RefKind.vbext_rk_Project)
                    {
                        var model = new VbaProjectReferenceModel(reference);
                        var vm = new RegisteredLibraryViewModel(model, vbe.ActiveVBProject);
                        _registeredLibraries.Add(vm);
                    }
                }
            }
        }
    }
}
