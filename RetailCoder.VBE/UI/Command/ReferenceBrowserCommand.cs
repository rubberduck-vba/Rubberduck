using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using NLog;
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
        private readonly IOpenFileDialog _filePicker;

        public ReferenceBrowserCommand(VBE vbe, RegisteredLibraryModelService service, IOpenFileDialog filePicker) 
            : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _service = service;
            _filePicker = filePicker;
        }

        protected override void ExecuteImpl(object parameter)
        {
            using (var vm = new ReferenceBrowserViewModel(_vbe, _service, _filePicker))
            using (var dialog = new ReferenceBrowserWindow(vm))
            {
                dialog.ShowDialog();
            }
        }
    }
}
