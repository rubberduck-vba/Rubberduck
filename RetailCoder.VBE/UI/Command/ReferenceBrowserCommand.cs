using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.UI.ReferenceBrowser;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that displays the Reference Browser window.
    /// </summary>
    [ComVisible(false)]
    public class ReferenceBrowserCommand : CommandBase
    {
        private readonly VBE _vbe;
        private readonly RegisteredLibraryModelService _service;

        public ReferenceBrowserCommand(VBE vbe, RegisteredLibraryModelService service)
        {
            _vbe = vbe;
            _service = service;
        }

        public override void Execute(object parameter)
        {
            var vm = new ReferenceBrowserViewModel(_vbe, _service);
            using (var dialog = new ReferenceBrowserWindow(vm))
            {
                dialog.ShowDialog();
            }
        }
    }
}
